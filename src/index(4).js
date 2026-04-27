/**
 * Cotizador SSO (Cloudflare Worker + D1) — Corporate MVP
 * - Entra ID (JWT) verification (RS256 JWKS)
 * - Admin: Tarifas, Descuentos, Matriz, Convenios
 * - User: Cotización (manual/bulk), Carga masiva + validación, Firma simple de convenios
 *
 * NOTE: This Worker is intentionally dependency-free (no npm packages required).
 */

export default {
  async fetch(req, env, ctx) {
    const rid = mkRid();
    const cors = corsHeaders(req, env);

    try {
      const url = new URL(req.url);

      // CORS preflight
      if (req.method === "OPTIONS") {
        return new Response("", { status: 204, headers: cors });
      }

      // Health (public)
      if (req.method === "GET" && url.pathname === "/health") {
        return j({ ok: true, service: "cotizador-sso-worker", rid, ts: new Date().toISOString() }, 200, cors);
      }

      // Lightweight rate limit (best-effort, per Worker isolate)
      const limit = num(env.RATE_LIMIT, 30);
      const windowMs = num(env.RATE_WINDOW_MS, 60_000);
      const rl = rateLimit(req, limit, windowMs);
      if (!rl.ok) {
        return j(
          { ok: false, rid, error: "RATE_LIMIT", message: `Demasiadas solicitudes. Intenta en ${rl.retryAfterSec}s.` },
          429,
          { ...cors, "Retry-After": String(rl.retryAfterSec) }
        );
      }

      // Route protection
      const isPublic = url.pathname === "/health";
      const isProtected = !isPublic;

      let claims = null;
      if (isProtected) {
        claims = await verifyEntraJwt(req, env, ctx);
      }

      // DB required for almost everything
      if (isProtected && !env.DB) {
        return j(
          { ok: false, rid, error: "DB_NOT_CONFIGURED", message: "Binding D1 (DB) no configurado en el Worker." },
          500,
          cors
        );
      }

      // Schema flags + RBAC context
      let schema = null;
      let authz = null;
      if (isProtected) {
        schema = await getSchemaState(env);
        authz = await getAuthzContext(env, claims, schema);
      }

      // Who am I
      if (req.method === "GET" && url.pathname === "/me") {
        return j(
          {
            ok: true,
            rid,
            rbac_enabled: !!schema?.rbac,
            schema: schema || {},
            user: {
              ...(safeUserFromClaims(claims) || {}),
              role: authz?.role || "advisor",
              supervisor_oid: authz?.user?.supervisor_oid || null,
              active: authz?.user?.active ?? 1,
              email: authz?.user?.email || (claims?.preferred_username || claims?.upn || null),
              display_name: authz?.user?.display_name || claims?.name || null,
            },
            permissions: {
              role: authz?.role || "advisor",
              admin: !!authz?.isAdmin,
              supervisor: !!authz?.isSupervisor,
              advisor: !!authz?.isAdvisor,
              can_manage_users: !!authz?.isAdmin,
              can_manage_pricing: !!authz?.isAdmin,
              can_approve_discounts: !!(authz?.isAdmin || authz?.isSupervisor),
              max_manual_discount_bps: authz?.max_manual_discount_bps ?? null,
              max_manual_discount_percent: authz?.max_manual_discount_bps != null ? (authz.max_manual_discount_bps / 100) : null,
            },
          },
          200,
          cors
        );
      }

      // ===== Admin: Usuarios (RBAC interno) =====
      if (url.pathname === "/admin/users" && req.method === "GET") {
        assertAdminRole(authz);
        if (!schema?.rbac) {
          return j({ ok: false, rid, error: "RBAC_NOT_ENABLED", message: "Tabla users no existe. Ejecuta la migración RBAC." }, 409, cors);
        }

        const q = (url.searchParams.get("q") || "").trim();
        const limitN = clampInt(url.searchParams.get("limit"), 50, 1, 300);

        const where = [];
        const binds = [];
        if (q) {
          where.push("(oid LIKE ? OR email LIKE ? OR display_name LIKE ?)");
          binds.push(`%${q}%`, `%${q}%`, `%${q}%`);
        }

        const sql = `
          SELECT oid, email, display_name, role, supervisor_oid, active, created_at, created_by, updated_at, updated_by
          FROM users
          ${where.length ? "WHERE " + where.join(" AND ") : ""}
          ORDER BY created_at DESC
          LIMIT ?
        `;
        binds.push(limitN);

        const rows = await env.DB.prepare(sql).bind(...binds).all();
        return j({ ok: true, rid, items: rows.results || [] }, 200, cors);
      }

      if (url.pathname === "/admin/users" && req.method === "POST") {
        assertAdminRole(authz);
        if (!schema?.rbac) {
          return j({ ok: false, rid, error: "RBAC_NOT_ENABLED", message: "Tabla users no existe. Ejecuta la migración RBAC." }, 409, cors);
        }
        const body = await readJson(req, num(env.MAX_BODY_BYTES, 80_000));
        const now = new Date().toISOString();

        const oid = String(body?.oid || "").trim();
        const email = body?.email ? String(body.email).trim() : null;
        const display_name = body?.display_name ? String(body.display_name).trim() : null;
        const role = normRole(body?.role || "advisor");
        let supervisor_oid = body?.supervisor_oid ? String(body.supervisor_oid).trim() : null;
        const active = body?.active == null ? 1 : (body.active ? 1 : 0);

        if (!oid) return j({ ok: false, rid, error: "BAD_OID", message: "oid es requerido" }, 400, cors);
        if (!role) return j({ ok: false, rid, error: "BAD_ROLE", message: "role inválido" }, 400, cors);
        if (role !== "advisor") supervisor_oid = null;
        if (supervisor_oid === oid) supervisor_oid = null;

        await env.DB.prepare(
          `INSERT INTO users (oid, email, display_name, role, supervisor_oid, active, created_at, created_by)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(oid, email, display_name, role, supervisor_oid, active, now, authz.user.oid).run();

        await audit(env, { actor: authz.user.oid, action: "USER_CREATE", entity: "users", entity_id: oid, before: null, after: { oid, email, display_name, role, supervisor_oid, active }, rid });

        return j({ ok: true, rid, created: true, oid }, 200, cors);
      }

      const mUser = url.pathname.match(/^\/admin\/users\/([^/]+)$/);
      if (mUser && (req.method === "PATCH" || req.method === "PUT")) {
        assertAdminRole(authz);
        if (!schema?.rbac) {
          return j({ ok: false, rid, error: "RBAC_NOT_ENABLED", message: "Tabla users no existe. Ejecuta la migración RBAC." }, 409, cors);
        }

        const oid = decodeURIComponent(mUser[1]);
        const before = await env.DB.prepare(`SELECT * FROM users WHERE oid = ?`).bind(oid).first();
        if (!before) return j({ ok: false, rid, error: "NOT_FOUND" }, 404, cors);

        const body = await readJson(req, num(env.MAX_BODY_BYTES, 80_000));
        const now = new Date().toISOString();

        const role = body?.role != null ? normRole(body.role) : before.role;
        let supervisor_oid = body?.supervisor_oid != null ? (String(body.supervisor_oid).trim() || null) : before.supervisor_oid;
        const active = body?.active == null ? before.active : (body.active ? 1 : 0);
        const email = body?.email != null ? (String(body.email).trim() || null) : before.email;
        const display_name = body?.display_name != null ? (String(body.display_name).trim() || null) : before.display_name;

        if (!role) return j({ ok: false, rid, error: "BAD_ROLE", message: "role inválido" }, 400, cors);
        if (role !== "advisor") supervisor_oid = null;
        if (supervisor_oid === oid) supervisor_oid = null;

        await env.DB.prepare(
          `UPDATE users
             SET email=?, display_name=?, role=?, supervisor_oid=?, active=?, updated_at=?, updated_by=?
           WHERE oid=?`
        ).bind(email, display_name, role, supervisor_oid, active, now, authz.user.oid, oid).run();

        await audit(env, { actor: authz.user.oid, action: "USER_UPDATE", entity: "users", entity_id: oid, before, after: { email, display_name, role, supervisor_oid, active }, rid });

        return j({ ok: true, rid, updated: true, oid }, 200, cors);
      }

      // ===== Aprobaciones de descuento =====
      if (url.pathname === "/approvals" && req.method === "GET") {
        assertSupervisorOrAdmin(authz);
        if (!schema?.approvals) {
          return j({ ok: false, rid, error: "APPROVALS_NOT_ENABLED", message: "Tabla quote_approvals no existe. Ejecuta la migración de workflow." }, 409, cors);
        }

        const status = String(url.searchParams.get("status") || "pending").trim();
        const limitN = clampInt(url.searchParams.get("limit"), 50, 1, 200);

        let rows;
        if (authz.isAdmin) {
          rows = await env.DB.prepare(
            `SELECT qa.id, qa.quote_id, qa.requested_by, qa.requested_at, qa.requested_discount_bps, qa.requested_discount_cents,
                    qa.reason, qa.status, qa.acted_by, qa.acted_at,
                    q.client_name, q.client_ruc, q.currency, q.subtotal_cents, q.discount_cents, q.total_cents, q.status AS quote_status,
                    u.display_name AS requested_by_name, u.email AS requested_by_email
             FROM quote_approvals qa
             JOIN quotes q ON q.id = qa.quote_id
             LEFT JOIN users u ON u.oid = qa.requested_by
             WHERE qa.status = ?
             ORDER BY qa.requested_at DESC
             LIMIT ?`
          ).bind(status, limitN).all();
        } else {
          // supervisor: only approvals from advisors in my team
          rows = await env.DB.prepare(
            `SELECT qa.id, qa.quote_id, qa.requested_by, qa.requested_at, qa.requested_discount_bps, qa.requested_discount_cents,
                    qa.reason, qa.status, qa.acted_by, qa.acted_at,
                    q.client_name, q.client_ruc, q.currency, q.subtotal_cents, q.discount_cents, q.total_cents, q.status AS quote_status,
                    u.display_name AS requested_by_name, u.email AS requested_by_email
             FROM quote_approvals qa
             JOIN quotes q ON q.id = qa.quote_id
             LEFT JOIN users u ON u.oid = qa.requested_by
             WHERE qa.status = ?
               AND u.role = 'advisor'
               AND u.supervisor_oid = ?
             ORDER BY qa.requested_at DESC
             LIMIT ?`
          ).bind(status, authz.user.oid, limitN).all();
        }

        return j({ ok: true, rid, items: rows.results || [] }, 200, cors);
      }

      const mApprovalAct = url.pathname.match(/^\/approvals\/(\d+)\/(approve|reject)$/);
      if (mApprovalAct && req.method === "POST") {
        assertSupervisorOrAdmin(authz);
        if (!schema?.approvals || !schema?.quotes_status) {
          return j({ ok: false, rid, error: "APPROVALS_NOT_ENABLED", message: "Workflow no habilitado. Ejecuta migración (quotes.status + quote_approvals)." }, 409, cors);
        }

        const id = Number(mApprovalAct[1]);
        const action = mApprovalAct[2];
        if (!id) return j({ ok: false, rid, error: "BAD_ID" }, 400, cors);

        const body = await readJson(req, num(env.MAX_BODY_BYTES, 30_000));
        const comment = String(body?.comment || "").trim() || null;

        // load approval + requester info for scope check
        const rec = await env.DB.prepare(
          `SELECT qa.*, u.role AS requester_role, u.supervisor_oid AS requester_supervisor_oid
           FROM quote_approvals qa
           LEFT JOIN users u ON u.oid = qa.requested_by
           WHERE qa.id = ?`
        ).bind(id).first();

        if (!rec) return j({ ok: false, rid, error: "NOT_FOUND" }, 404, cors);
        if (String(rec.status) !== "pending") return j({ ok: false, rid, error: "NOT_PENDING", message: "La solicitud ya fue procesada." }, 409, cors);

        if (!authz.isAdmin) {
          const ok = rec.requester_role === "advisor" && rec.requester_supervisor_oid === authz.user.oid;
          if (!ok) {
            const e = new Error("FORBIDDEN");
            e.status = 403;
            throw e;
          }
        }

        const now = new Date().toISOString();
        const acted_by = authz.user.oid;
        const approvalStatus = action === "approve" ? "approved" : "rejected";
        const quoteStatus = action === "approve" ? "approved" : "draft";

        await env.DB.prepare(
          `UPDATE quote_approvals
             SET status=?, acted_by=?, acted_at=?, acted_comment=?
           WHERE id=?`
        ).bind(approvalStatus, acted_by, now, comment, id).run();

        await setQuoteStatus(env, schema, rec.quote_id, {
          status: quoteStatus,
          updated_at: now,
          updated_by: acted_by,
          approved_at: action === "approve" ? now : null,
          approved_by: action === "approve" ? acted_by : null,
        });

        await audit(env, { actor: acted_by, action: action === "approve" ? "QUOTE_APPROVE" : "QUOTE_REJECT", entity: "quote_approvals", entity_id: String(id), before: rec, after: { status: approvalStatus, quote_status: quoteStatus, comment }, rid });

        return j({ ok: true, rid, id, status: approvalStatus, quote_id: rec.quote_id, quote_status: quoteStatus }, 200, cors);
      }

      // ===== Admin: Tarifas =====
      if (url.pathname === "/admin/tariffs" && req.method === "GET") {
        assertAdminRole(authz);

        const qService = (url.searchParams.get("service_code") || "").trim();
        const qExam = (url.searchParams.get("exam_type") || "").trim();
        const qActive = (url.searchParams.get("active") || "").trim();
        const limitN = clampInt(url.searchParams.get("limit"), 200, 1, 500);

        const where = [];
        const binds = [];
        if (qService) {
          where.push("service_code LIKE ?");
          binds.push(`%${qService}%`);
        }
        if (qExam) {
          where.push("exam_type LIKE ?");
          binds.push(`%${qExam}%`);
        }
        if (qActive === "0" || qActive === "1") {
          where.push("active = ?");
          binds.push(Number(qActive));
        }

        const sql = `
          SELECT id, service_code, service_name, exam_type, sex, min_age, max_age, currency,
                 price_cents, effective_from, effective_to, version, active, created_at, created_by
          FROM tariffs
          ${where.length ? "WHERE " + where.join(" AND ") : ""}
          ORDER BY created_at DESC, id DESC
          LIMIT ?
        `;
        binds.push(limitN);

        const rows = await env.DB.prepare(sql).bind(...binds).all();
        return j({ ok: true, rid, items: rows.results || [] }, 200, cors);
      }

      if (url.pathname === "/admin/tariffs/import" && req.method === "POST") {
        assertAdminRole(authz);

        const body = await readJson(req, num(env.MAX_BODY_BYTES, 250_000));
        const version = String(body?.version || "").trim() || new Date().toISOString().slice(0, 10);
        const effective_from = toIsoDate(String(body?.effective_from || "").trim() || new Date().toISOString().slice(0, 10));
        const effective_to = body?.effective_to ? toIsoDate(String(body.effective_to).trim()) : null;

        const rows = Array.isArray(body?.rows) ? body.rows : [];
        if (!rows.length) return j({ ok: false, rid, error: "NO_ROWS", message: "rows debe ser un array no vacío." }, 400, cors);

        // Insert
        const actor = actorId(claims);
        const now = new Date().toISOString();

        const stmts = [];
        for (const r of rows) {
          const service_code = normCode(r?.service_code);
          const service_name = String(r?.service_name || "").trim() || null;
          const exam_type = normText(r?.exam_type) || "GENERAL";
          const sex = normSex(r?.sex);
          const min_age = clampInt(r?.min_age, 0, 0, 120);
          const max_age = clampInt(r?.max_age, 200, 0, 200);
          const currency = normText(r?.currency) || "USD";
          const price_cents = moneyToCents(r?.price);

          if (!service_code) {
            return j({ ok: false, rid, error: "BAD_ROW", message: "service_code es requerido en todas las filas." }, 400, cors);
          }
          if (!Number.isFinite(price_cents) || price_cents < 0) {
            return j({ ok: false, rid, error: "BAD_ROW", message: `price inválido para ${service_code}.` }, 400, cors);
          }

          stmts.push(
            env.DB.prepare(
              `INSERT INTO tariffs
                (service_code, service_name, exam_type, sex, min_age, max_age, currency, price_cents,
                 effective_from, effective_to, version, active, created_at, created_by)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?)`
            ).bind(
              service_code,
              service_name,
              exam_type,
              sex,
              min_age,
              max_age,
              currency,
              price_cents,
              effective_from,
              effective_to,
              version,
              now,
              actor
            )
          );
        }

        await env.DB.batch(stmts);

        await audit(env, {
          actor,
          action: "TARIFFS_IMPORT",
          entity: "tariffs",
          entity_id: version,
          before: null,
          after: { version, effective_from, effective_to, count: rows.length },
          rid,
        });

        return j({ ok: true, rid, imported: rows.length, version }, 200, cors);
      }

      // ===== Admin: Descuentos =====
      if (url.pathname === "/admin/discounts" && req.method === "GET") {
        assertAdminRole(authz);

        const limitN = clampInt(url.searchParams.get("limit"), 200, 1, 500);
        const rows = await env.DB.prepare(
          `SELECT id, name, priority, type, value_bps, value_cents, currency, stackable,
                  conditions_json, active, effective_from, effective_to, created_at, created_by
           FROM discounts
           ORDER BY priority ASC, id DESC
           LIMIT ?`
        ).bind(limitN).all();

        const items = (rows.results || []).map(d => ({
          ...d,
          conditions: safeJsonParse(d.conditions_json) || {},
        }));
        return j({ ok: true, rid, items }, 200, cors);
      }

      if (url.pathname === "/admin/discounts" && req.method === "POST") {
        assertAdminRole(authz);

        const body = await readJson(req, num(env.MAX_BODY_BYTES, 80_000));
        const actor = actorId(claims);
        const now = new Date().toISOString();

        const rec = normalizeDiscount(body);
        if (!rec.ok) return j({ ok: false, rid, error: rec.error, message: rec.message }, 400, cors);

        const stmt = env.DB.prepare(
          `INSERT INTO discounts
            (name, priority, type, value_bps, value_cents, currency, stackable, conditions_json, active, effective_from, effective_to, created_at, created_by)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(
          rec.value.name,
          rec.value.priority,
          rec.value.type,
          rec.value.value_bps,
          rec.value.value_cents,
          rec.value.currency,
          rec.value.stackable,
          JSON.stringify(rec.value.conditions || {}),
          rec.value.active,
          rec.value.effective_from,
          rec.value.effective_to,
          now,
          actor
        );

        const out = await stmt.run();
        await audit(env, { actor, action: "DISCOUNT_CREATE", entity: "discounts", entity_id: String(out.meta?.last_row_id || ""), before: null, after: rec.value, rid });

        return j({ ok: true, rid, created: true, id: out.meta?.last_row_id || null }, 200, cors);
      }

      // Update discount
      const mDiscount = url.pathname.match(/^\/admin\/discounts\/(\d+)$/);
      if (mDiscount && (req.method === "PUT" || req.method === "PATCH")) {
        assertAdminRole(authz);

        const id = Number(mDiscount[1]);
        if (!id) return j({ ok: false, rid, error: "BAD_ID" }, 400, cors);

        const beforeRow = await env.DB.prepare(`SELECT * FROM discounts WHERE id = ?`).bind(id).first();
        if (!beforeRow) return j({ ok: false, rid, error: "NOT_FOUND" }, 404, cors);

        const body = await readJson(req, num(env.MAX_BODY_BYTES, 80_000));
        const actor = actorId(claims);

        const rec = normalizeDiscount({ ...beforeRow, ...body }, true);
        if (!rec.ok) return j({ ok: false, rid, error: rec.error, message: rec.message }, 400, cors);

        await env.DB.prepare(
          `UPDATE discounts
             SET name=?, priority=?, type=?, value_bps=?, value_cents=?, currency=?, stackable=?, conditions_json=?, active=?, effective_from=?, effective_to=?
           WHERE id=?`
        ).bind(
          rec.value.name,
          rec.value.priority,
          rec.value.type,
          rec.value.value_bps,
          rec.value.value_cents,
          rec.value.currency,
          rec.value.stackable,
          JSON.stringify(rec.value.conditions || {}),
          rec.value.active,
          rec.value.effective_from,
          rec.value.effective_to,
          id
        ).run();

        await audit(env, { actor, action: "DISCOUNT_UPDATE", entity: "discounts", entity_id: String(id), before: beforeRow, after: rec.value, rid });

        return j({ ok: true, rid, updated: true, id }, 200, cors);
      }

      // ===== Admin: Matriz =====
      if (url.pathname === "/admin/matrix" && req.method === "GET") {
        assertAdminRole(authz);

        const limitN = clampInt(url.searchParams.get("limit"), 200, 1, 500);
        const rows = await env.DB.prepare(
          `SELECT id, name, exam_type, sex, min_age, max_age, services_json, notes, active, created_at, created_by
           FROM matrix_rules
           ORDER BY created_at DESC, id DESC
           LIMIT ?`
        ).bind(limitN).all();

        const items = (rows.results || []).map(x => ({ ...x, services: safeJsonParse(x.services_json) || [] }));
        return j({ ok: true, rid, items }, 200, cors);
      }

      if (url.pathname === "/admin/matrix" && req.method === "POST") {
        assertAdminRole(authz);

        const body = await readJson(req, num(env.MAX_BODY_BYTES, 80_000));
        const actor = actorId(claims);
        const now = new Date().toISOString();

        const rec = normalizeMatrixRule(body);
        if (!rec.ok) return j({ ok: false, rid, error: rec.error, message: rec.message }, 400, cors);

        const out = await env.DB.prepare(
          `INSERT INTO matrix_rules
            (name, exam_type, sex, min_age, max_age, services_json, notes, active, created_at, created_by)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(
          rec.value.name,
          rec.value.exam_type,
          rec.value.sex,
          rec.value.min_age,
          rec.value.max_age,
          JSON.stringify(rec.value.services || []),
          rec.value.notes,
          rec.value.active,
          now,
          actor
        ).run();

        await audit(env, { actor, action: "MATRIX_CREATE", entity: "matrix_rules", entity_id: String(out.meta?.last_row_id || ""), before: null, after: rec.value, rid });

        return j({ ok: true, rid, created: true, id: out.meta?.last_row_id || null }, 200, cors);
      }

      const mMatrix = url.pathname.match(/^\/admin\/matrix\/(\d+)$/);
      if (mMatrix && (req.method === "PUT" || req.method === "PATCH" || req.method === "DELETE")) {
        assertAdminRole(authz);
        const id = Number(mMatrix[1]);
        if (!id) return j({ ok: false, rid, error: "BAD_ID" }, 400, cors);

        const beforeRow = await env.DB.prepare(`SELECT * FROM matrix_rules WHERE id = ?`).bind(id).first();
        if (!beforeRow) return j({ ok: false, rid, error: "NOT_FOUND" }, 404, cors);

        const actor = actorId(claims);

        if (req.method === "DELETE") {
          await env.DB.prepare(`DELETE FROM matrix_rules WHERE id = ?`).bind(id).run();
          await audit(env, { actor, action: "MATRIX_DELETE", entity: "matrix_rules", entity_id: String(id), before: beforeRow, after: null, rid });
          return j({ ok: true, rid, deleted: true, id }, 200, cors);
        }

        const body = await readJson(req, num(env.MAX_BODY_BYTES, 80_000));
        const rec = normalizeMatrixRule({ ...beforeRow, ...body }, true);
        if (!rec.ok) return j({ ok: false, rid, error: rec.error, message: rec.message }, 400, cors);

        await env.DB.prepare(
          `UPDATE matrix_rules
             SET name=?, exam_type=?, sex=?, min_age=?, max_age=?, services_json=?, notes=?, active=?
           WHERE id=?`
        ).bind(
          rec.value.name,
          rec.value.exam_type,
          rec.value.sex,
          rec.value.min_age,
          rec.value.max_age,
          JSON.stringify(rec.value.services || []),
          rec.value.notes,
          rec.value.active,
          id
        ).run();

        await audit(env, { actor, action: "MATRIX_UPDATE", entity: "matrix_rules", entity_id: String(id), before: beforeRow, after: rec.value, rid });

        return j({ ok: true, rid, updated: true, id }, 200, cors);
      }

      // ===== Convenios (Agreements) =====
      // List (auth)
      if (url.pathname === "/agreements" && req.method === "GET") {
        const limitN = clampInt(url.searchParams.get("limit"), 50, 1, 200);
        const rows = await env.DB.prepare(
          `SELECT id, client_name, client_ruc, title, status, effective_from, effective_to, created_at, created_by
           FROM agreements
           ORDER BY created_at DESC
           LIMIT ?`
        ).bind(limitN).all();

        return j({ ok: true, rid, items: rows.results || [] }, 200, cors);
      }

      // Create (admin)
      if (url.pathname === "/agreements" && req.method === "POST") {
        assertAdminRole(authz);
        const body = await readJson(req, num(env.MAX_BODY_BYTES, 80_000));

        const actor = actorId(claims);
        const now = new Date().toISOString();

        const id = `agr_${cryptoRandomId()}`;
        const client_name = String(body?.client_name || "").trim();
        const client_ruc = String(body?.client_ruc || "").trim();
        const title = String(body?.title || "").trim() || "Convenio";
        const body_text = String(body?.body_text || "").trim() || "";
        const status = String(body?.status || "draft").trim();
        const effective_from = toIsoDate(String(body?.effective_from || new Date().toISOString().slice(0,10)));
        const effective_to = body?.effective_to ? toIsoDate(String(body?.effective_to)) : null;

        if (!client_name) return j({ ok: false, rid, error: "BAD_CLIENT", message: "client_name es requerido." }, 400, cors);
        if (!client_ruc) return j({ ok: false, rid, error: "BAD_CLIENT", message: "client_ruc es requerido." }, 400, cors);

        await env.DB.prepare(
          `INSERT INTO agreements
            (id, client_name, client_ruc, title, body_text, status, effective_from, effective_to, created_at, created_by)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(id, client_name, client_ruc, title, body_text, status, effective_from, effective_to, now, actor).run();

        await audit(env, { actor, action: "AGREEMENT_CREATE", entity: "agreements", entity_id: id, before: null, after: { client_name, client_ruc, title, status }, rid });

        return j({ ok: true, rid, id }, 200, cors);
      }

      // Get agreement detail
      const mAgr = url.pathname.match(/^\/agreements\/(agr_[A-Za-z0-9\-_]+)$/);
      if (mAgr && req.method === "GET") {
        const id = mAgr[1];
        const a = await env.DB.prepare(`SELECT * FROM agreements WHERE id = ?`).bind(id).first();
        if (!a) return j({ ok: false, rid, error: "NOT_FOUND" }, 404, cors);

        const sigs = await env.DB.prepare(
          `SELECT id, signer_name, signer_email, signer_oid, signed_at, signature_hash
           FROM agreement_signatures
           WHERE agreement_id = ?
           ORDER BY signed_at DESC`
        ).bind(id).all();

        return j({ ok: true, rid, agreement: a, signatures: sigs.results || [] }, 200, cors);
      }

      // Sign agreement (auth)
      const mSign = url.pathname.match(/^\/agreements\/(agr_[A-Za-z0-9\-_]+)\/sign$/);
      if (mSign && req.method === "POST") {
        const agreement_id = mSign[1];
        const a = await env.DB.prepare(`SELECT id, status FROM agreements WHERE id = ?`).bind(agreement_id).first();
        if (!a) return j({ ok: false, rid, error: "NOT_FOUND" }, 404, cors);

        const body = await readJson(req, num(env.MAX_BODY_BYTES, 30_000));

        const signer_name = String(body?.signer_name || claims?.name || "").trim();
        const signer_email = String(body?.signer_email || claims?.preferred_username || claims?.upn || "").trim();
        const signature_text = String(body?.signature_text || "").trim();

        if (!signer_name) return j({ ok: false, rid, error: "BAD_SIGNER", message: "signer_name es requerido." }, 400, cors);
        if (!signature_text) return j({ ok: false, rid, error: "BAD_SIGNATURE", message: "signature_text es requerido." }, 400, cors);

        const signed_at = new Date().toISOString();
        const ip = req.headers.get("CF-Connecting-IP") || "";
        const ua = req.headers.get("User-Agent") || "";

        const signer_oid = String(claims?.oid || "");
        const hash = await sha256Hex(`${agreement_id}|${signer_oid}|${signer_email}|${signature_text}|${signed_at}`);

        await env.DB.prepare(
          `INSERT INTO agreement_signatures
            (agreement_id, signer_name, signer_email, signer_oid, signed_at, signature_text, signature_hash, ip, user_agent)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(agreement_id, signer_name, signer_email, signer_oid, signed_at, signature_text, hash, ip, ua).run();

        await audit(env, { actor: actorId(claims), action: "AGREEMENT_SIGN", entity: "agreements", entity_id: agreement_id, before: null, after: { signer_name, signer_email, signature_hash: hash }, rid });

        return j({ ok: true, rid, signed: true, signature_hash: hash, signed_at }, 200, cors);
      }

      // ===== Carga masiva: validación =====
      if (url.pathname === "/bulk/validate" && req.method === "POST") {
        const body = await readJson(req, num(env.MAX_BODY_BYTES, 400_000));
        const rows = Array.isArray(body?.rows) ? body.rows : [];
        if (!rows.length) return j({ ok: false, rid, error: "NO_ROWS", message: "rows debe ser un array no vacío." }, 400, cors);

        const asOf = body?.as_of ? toIsoDate(String(body.as_of)) : new Date().toISOString().slice(0, 10);

        const out = validateBulkRows(rows, asOf);
        return j({ ok: true, rid, as_of: asOf, ...out }, 200, cors);
      }

      // ===== Cotizaciones =====
      if (url.pathname === "/quotes/preview" && req.method === "POST") {
        const body = await readJson(req, num(env.MAX_BODY_BYTES, 400_000));
        const asOf = body?.valid_on ? toIsoDate(String(body.valid_on)) : new Date().toISOString().slice(0, 10);

        const result = await computeQuote(env, rid, claims, authz, schema, { ...body, valid_on: asOf }, { save: false });
        return j({ ok: true, rid, quote: result }, 200, cors);
      }

      if (url.pathname === "/quotes" && req.method === "POST") {
        const body = await readJson(req, num(env.MAX_BODY_BYTES, 400_000));
        const asOf = body?.valid_on ? toIsoDate(String(body.valid_on)) : new Date().toISOString().slice(0, 10);

        const result = await computeQuote(env, rid, claims, authz, schema, { ...body, valid_on: asOf }, { save: true });
        return j({ ok: true, rid, quote: result }, 200, cors);
      }

      if (url.pathname === "/quotes" && req.method === "GET") {
        const limitN = clampInt(url.searchParams.get("limit"), 30, 1, 100);
        const q = (url.searchParams.get("q") || "").trim();
        const statusFilter = (url.searchParams.get("status") || "").trim();

        const where = [];
        const binds = [];

        // RBAC: advisors only their quotes; supervisors their team; admins all
        if (!authz?.isAdmin) {
          if (authz?.isSupervisor && schema?.rbac) {
            const team = await env.DB.prepare(
              `SELECT oid FROM users WHERE supervisor_oid = ? AND active = 1`
            ).bind(authz.user.oid).all();
            const oids = [authz.user.oid, ...((team.results || []).map(r => r.oid).filter(Boolean))];
            where.push(`created_by IN (${oids.map(() => "?").join(",")})`);
            binds.push(...oids);
          } else {
            where.push("created_by = ?");
            binds.push(authz.user.oid);
          }
        }

        if (q) {
          where.push("(id LIKE ? OR client_name LIKE ? OR client_ruc LIKE ?)");
          binds.push(`%${q}%`, `%${q}%`, `%${q}%`);
        }

        if (schema?.quotes_status && statusFilter) {
          where.push("status = ?");
          binds.push(statusFilter);
        }

        const cols = [
          "id",
          "created_at",
          "created_by",
          "client_name",
          "client_ruc",
          "valid_on",
          "mode",
          "currency",
          "subtotal_cents",
          "discount_cents",
          "total_cents",
        ];

        if (schema?.quotes_status) {
          cols.push("status", "updated_at", "updated_by", "approved_at", "approved_by");
        } else {
          cols.push("'draft' AS status");
        }

        if (schema?.approvals) {
          cols.push(
            "(SELECT qa.status FROM quote_approvals qa WHERE qa.quote_id = quotes.id ORDER BY qa.requested_at DESC LIMIT 1) AS approval_status"
          );
        }

        const sql = `
          SELECT ${cols.join(", ")}
          FROM quotes
          ${where.length ? "WHERE " + where.join(" AND ") : ""}
          ORDER BY created_at DESC
          LIMIT ?
        `;
        binds.push(limitN);

        const rows = await env.DB.prepare(sql).bind(...binds).all();
        return j({ ok: true, rid, items: rows.results || [] }, 200, cors);
      }

      // Update quote status (sent|won|lost|archived)
      const mQuoteStatus = url.pathname.match(/^\/quotes\/(q_[A-Za-z0-9\-_]+)\/status$/);
      if (mQuoteStatus && (req.method === "PATCH" || req.method === "POST")) {
        if (!schema?.quotes_status) {
          return j({ ok: false, rid, error: "WORKFLOW_NOT_ENABLED", message: "quotes.status no existe. Ejecuta migración." }, 409, cors);
        }
        const id = mQuoteStatus[1];
        const row = await env.DB.prepare(`SELECT id, created_by, status, result_json FROM quotes WHERE id = ?`).bind(id).first();
        if (!row) return j({ ok: false, rid, error: "NOT_FOUND" }, 404, cors);

        await assertCanAccessQuote(env, schema, authz, row);

        const body = await readJson(req, num(env.MAX_BODY_BYTES, 30_000));
        const status = String(body?.status || "").trim();
        const allowed = new Set(["draft", "pending_approval", "approved", "sent", "won", "lost", "archived"]);
        if (!allowed.has(status)) return j({ ok: false, rid, error: "BAD_STATUS", message: "status inválido" }, 400, cors);

        // Optional: prevent sending if pending approval
        if (status === "sent" && String(row.status) === "pending_approval") {
          return j({ ok: false, rid, error: "PENDING_APPROVAL", message: "La cotización está pendiente de aprobación." }, 409, cors);
        }

        const now = new Date().toISOString();
        const actor = authz.user.oid;
        await setQuoteStatus(env, schema, id, { status, updated_at: now, updated_by: actor });
        await audit(env, { actor, action: "QUOTE_STATUS", entity: "quotes", entity_id: id, before: { status: row.status }, after: { status }, rid });

        return j({ ok: true, rid, id, status }, 200, cors);
      }

      const mQuote = url.pathname.match(/^\/quotes\/(q_[A-Za-z0-9\-_]+)$/);
      if (mQuote && req.method === "GET") {
        const id = mQuote[1];
        const row = await env.DB.prepare(`SELECT * FROM quotes WHERE id = ?`).bind(id).first();
        if (!row) return j({ ok: false, rid, error: "NOT_FOUND" }, 404, cors);

        await assertCanAccessQuote(env, schema, authz, row);

        return j(
          {
            ok: true,
            rid,
            quote: {
              ...row,
              payload: safeJsonParse(row.payload_json) || null,
              result: safeJsonParse(row.result_json) || null,
            },
          },
          200,
          cors
        );
      }

      // Not found
      return j({ ok: false, rid, error: "NOT_FOUND" }, 404, cors);
    } catch (e) {
      const msg = String(e?.message || e);
      const status =
        Number(e?.status) ||
        (msg.startsWith("BODY_TOO_LARGE") ? 413 : 500);

      const code =
        msg === "BAD_JSON" ? "BAD_JSON" :
        (msg.startsWith("BODY_TOO_LARGE") ? "BODY_TOO_LARGE" :
        (status === 401 ? "UNAUTHORIZED" :
        (status === 403 ? "FORBIDDEN" : "WORKER_ERROR")));

      return j(
        {
          ok: false,
          rid,
          error: code,
          message: msg,
          stack: status >= 500 && e?.stack ? String(e.stack).slice(0, 1200) : null,
        },
        status,
        cors
      );
    }
  },
};

// ======================
// Quote computation
// ======================
async function computeQuote(env, rid, claims, authz, schema, input, { save }) {
  const mode = String(input?.mode || "manual").trim().toLowerCase();
  const valid_on = toIsoDate(String(input?.valid_on || new Date().toISOString().slice(0, 10)));
  const exam_type = normText(input?.exam_type) || "GENERAL";

  const client_name = String(input?.client?.name || input?.client_name || "").trim();
  const client_ruc = String(input?.client?.ruc || input?.client_ruc || "").trim();
  const contact_name = String(input?.client?.contact_name || input?.contact_name || "").trim() || null;
  const contact_email = String(input?.client?.contact_email || input?.contact_email || "").trim() || null;

  if (!client_name) throw new Error("client.name es requerido");
  if (!client_ruc) throw new Error("client.ruc es requerido");

  const warnings = [];

  // Build pricing lines
  let lines = [];
  let peopleCount = 0;

  if (mode === "bulk") {
    const rows = Array.isArray(input?.rows) ? input.rows : [];
    if (!rows.length) throw new Error("rows es requerido para modo bulk");

    // validate + normalize
    const validated = validateBulkRows(rows, valid_on);
    peopleCount = validated.valid_rows.length;
    if (!peopleCount) {
      return {
        id: null,
        mode,
        valid_on,
        exam_type,
        client: { name: client_name, ruc: client_ruc, contact_name, contact_email },
        currency: "USD",
        items: [],
        subtotal_cents: 0,
        discounts: [],
        discount_total_cents: 0,
        total_cents: 0,
        warnings: ["No hay filas válidas. Revisa errores de carga masiva."].concat(validated.errors.slice(0, 10).map(e => `Fila ${e.row}: ${e.messages.join("; ")}`)),
      };
    }

    // Load matrix rules for exam_type
    const rules = await loadMatrixRules(env, exam_type);

    // For each employee, determine services and tariff per service
    const perKey = new Map(); // key => aggregated line
    for (const emp of validated.valid_rows) {
      const sex = emp.sex;
      const age = emp.age;

      const services = pickServicesForEmployee(rules, { sex, age });
      if (!services.length) {
        warnings.push(`Sin regla de matriz para ${emp.full_name || emp.id_number} (edad ${age}, sexo ${sex}).`);
        continue;
      }

      for (const service_code of services) {
        const t = await findTariff(env, { service_code, exam_type, sex, age, valid_on });
        if (!t) {
          warnings.push(`Sin tarifa para ${service_code} (${exam_type}) para sexo ${sex} edad ${age}.`);
          continue;
        }
        const key = `${service_code}::${t.price_cents}::${t.currency}::${t.exam_type}::${t.sex}::${t.min_age}-${t.max_age}`;
        const existing = perKey.get(key) || {
          service_code,
          service_name: t.service_name || null,
          qty: 0,
          unit_price_cents: t.price_cents,
          currency: t.currency || "USD",
          exam_type: t.exam_type,
          meta: { sex: t.sex, min_age: t.min_age, max_age: t.max_age },
        };
        existing.qty += 1;
        perKey.set(key, existing);
      }
    }

    lines = [...perKey.values()].map(x => ({
      service_code: x.service_code,
      service_name: x.service_name,
      qty: x.qty,
      unit_price_cents: x.unit_price_cents,
      line_total_cents: x.qty * x.unit_price_cents,
      currency: x.currency,
      meta: x.meta,
    }));
  } else {
    // manual mode
    const items = Array.isArray(input?.items) ? input.items : [];
    if (!items.length) throw new Error("items es requerido para modo manual");

    const sex = normSex(input?.sex || "A");
    const age = clampInt(input?.age, 35, 0, 200);

    const out = [];
    for (const it of items) {
      const service_code = normCode(it?.service_code);
      const qty = clampInt(it?.qty, 1, 1, 10_000);
      if (!service_code) continue;

      // First try strict matching (sex/age), then fallback to sex ANY (common corporate pricing)
      let t = await findTariff(env, { service_code, exam_type, sex, age, valid_on });
      if (!t) t = await findTariff(env, { service_code, exam_type, sex: "A", age: 35, valid_on });

      if (!t) {
        warnings.push(`Sin tarifa para ${service_code} (${exam_type}).`);
        continue;
      }

      out.push({
        service_code,
        service_name: t.service_name || null,
        qty,
        unit_price_cents: t.price_cents,
        line_total_cents: qty * t.price_cents,
        currency: t.currency || "USD",
        meta: { sex: t.sex, min_age: t.min_age, max_age: t.max_age },
      });
    }
    lines = out;
  }

  // currency: assume single currency
  const currency = (lines[0]?.currency) || "USD";
  const subtotal_cents = lines.reduce((s, x) => s + (x.line_total_cents || 0), 0);

  // Apply discounts
  const discounts = await loadDiscounts(env, valid_on);
  const applied = [];
  let discount_total_cents = 0;

  // base info for conditions
  const totalQty = lines.reduce((s, x) => s + (x.qty || 0), 0);

  for (const d of discounts) {
    const cond = d.conditions || {};
    const check = checkDiscountConditions(cond, { exam_type, client_ruc, subtotal_cents, totalQty, peopleCount, lines });
    if (!check.ok) continue;

    const baseCents = check.base_cents;
    if (baseCents <= 0) continue;

    let amount = 0;
    if (d.type === "percent") {
      amount = Math.round(baseCents * (d.value_bps / 10_000));
    } else {
      amount = Math.min(baseCents, d.value_cents);
    }
    if (amount <= 0) continue;

    applied.push({
      id: d.id,
      name: d.name,
      type: d.type,
      amount_cents: amount,
      stackable: !!d.stackable,
      reason: check.reason || "Condiciones cumplidas",
    });

    discount_total_cents += amount;

    if (!d.stackable) break;
  }

  let total_cents = Math.max(0, subtotal_cents - discount_total_cents);

  // Manual discount (workflow)
  const manual_percent_raw = input?.manual_discount_percent ?? input?.discount_percent ?? null;
  const manual_reason_raw = input?.manual_discount_reason ?? input?.discount_reason ?? null;
  const manual_reason = manual_reason_raw != null ? String(manual_reason_raw).trim() : "";

  let manual_discount_bps = 0;
  if (manual_percent_raw != null && String(manual_percent_raw).trim() !== "") {
    const p = Number(String(manual_percent_raw).replace(",", "."));
    if (Number.isFinite(p) && p > 0) {
      manual_discount_bps = Math.max(0, Math.min(10_000, Math.round(p * 100)));
    }
  }

  let manual_discount_cents = 0;
  if (manual_discount_bps > 0) {
    manual_discount_cents = Math.round(subtotal_cents * (manual_discount_bps / 10_000));
    // never discount below 0
    manual_discount_cents = Math.max(0, Math.min(manual_discount_cents, Math.max(0, subtotal_cents - discount_total_cents)));
    if (manual_discount_cents > 0) {
      applied.push({
        id: "manual",
        name: "Descuento manual",
        type: "percent",
        value_bps: manual_discount_bps,
        amount_cents: manual_discount_cents,
        stackable: true,
        reason: manual_reason || "Descuento manual",
        manual: true,
      });
      discount_total_cents += manual_discount_cents;
    }
  }

  // Recompute totals after manual discount
  total_cents = Math.max(0, subtotal_cents - discount_total_cents);

  const maxBps = authz?.max_manual_discount_bps;
  const needsApproval = !!(manual_discount_bps > 0 && Number.isFinite(maxBps) && manual_discount_bps > maxBps && !authz?.isAdmin);
  const approval_target_role = needsApproval
    ? (authz?.role === "advisor" ? (authz?.user?.supervisor_oid ? "supervisor" : "admin") : "admin")
    : null;
  const status = schema?.quotes_status ? (needsApproval ? "pending_approval" : "draft") : "draft";

  const quote = {
    id: null,
    mode,
    valid_on,
    exam_type,
    client: { name: client_name, ruc: client_ruc, contact_name, contact_email },
    status,
    currency,
    items: lines,
    subtotal_cents,
    discounts: applied,
    discount_total_cents,
    total_cents,
    warnings: warnings.slice(0, 50),
    workflow: {
      manual_discount_bps: manual_discount_bps || 0,
      manual_discount_percent: manual_discount_bps ? manual_discount_bps / 100 : 0,
      manual_discount_cents: manual_discount_cents || 0,
      max_manual_discount_bps: Number.isFinite(maxBps) ? maxBps : null,
      max_manual_discount_percent: Number.isFinite(maxBps) ? maxBps / 100 : null,
      approval_required: needsApproval,
      approval_target_role,
      approval_target_oid: approval_target_role === "supervisor" ? (authz?.user?.supervisor_oid || null) : null,
      approval_id: null,
    },
    meta: {
      peopleCount: peopleCount || null,
      created_by: authz?.user?.oid || actorId(claims),
    },
  };

  if (save) {
    if (needsApproval && (!schema?.approvals || !schema?.quotes_status)) {
      const e = new Error("APPROVALS_NOT_ENABLED");
      e.status = 409;
      throw e;
    }

    const id = `q_${cryptoRandomId()}`;
    const now = new Date().toISOString();
    const actor = authz?.user?.oid || actorId(claims);

    quote.id = id;
    quote.created_at = now;
    quote.meta.created_by = actor;

    if (schema?.quotes_status) {
      await env.DB.prepare(
        `INSERT INTO quotes
          (id, created_at, created_by, client_name, client_ruc, contact_name, contact_email, valid_on, mode,
           currency, subtotal_cents, discount_cents, total_cents,
           status, updated_at, updated_by, approved_at, approved_by,
           payload_json, result_json)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(
        id,
        now,
        actor,
        client_name,
        client_ruc,
        contact_name,
        contact_email,
        valid_on,
        mode,
        currency,
        subtotal_cents,
        discount_total_cents,
        total_cents,
        status,
        now,
        actor,
        null,
        null,
        JSON.stringify(input || {}),
        JSON.stringify(quote || {})
      ).run();
    } else {
      await env.DB.prepare(
        `INSERT INTO quotes
          (id, created_at, created_by, client_name, client_ruc, contact_name, contact_email, valid_on, mode,
           currency, subtotal_cents, discount_cents, total_cents, payload_json, result_json)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(
        id,
        now,
        actor,
        client_name,
        client_ruc,
        contact_name,
        contact_email,
        valid_on,
        mode,
        currency,
        subtotal_cents,
        discount_total_cents,
        total_cents,
        JSON.stringify(input || {}),
        JSON.stringify(quote || {})
      ).run();
    }

    let approval_id = null;
    if (needsApproval && schema?.approvals) {
      const out = await env.DB.prepare(
        `INSERT INTO quote_approvals
          (quote_id, requested_by, requested_at, requested_discount_bps, requested_discount_cents, reason, status)
         VALUES (?, ?, ?, ?, ?, ?, 'pending')`
      ).bind(id, actor, now, manual_discount_bps, manual_discount_cents, manual_reason || null).run();
      approval_id = out.meta?.last_row_id || null;
      quote.workflow.approval_id = approval_id;
    }

    await audit(env, { actor, action: "QUOTE_CREATE", entity: "quotes", entity_id: id, before: null, after: { client_ruc, subtotal_cents, total_cents, status }, rid });
    if (approval_id) {
      await audit(env, { actor, action: "QUOTE_APPROVAL_REQUEST", entity: "quote_approvals", entity_id: String(approval_id), before: null, after: { quote_id: id, requested_discount_bps: manual_discount_bps, requested_discount_cents: manual_discount_cents }, rid });
    }
  }

  return quote;
}

async function findTariff(env, { service_code, exam_type, sex, age, valid_on }) {
  const ex = normText(exam_type) || "GENERAL";
  const sx = normSex(sex);
  const a = clampInt(age, 35, 0, 200);
  const d = toIsoDate(valid_on);

  const row = await env.DB.prepare(
    `SELECT id, service_code, service_name, exam_type, sex, min_age, max_age, currency, price_cents, effective_from, effective_to, version
     FROM tariffs
     WHERE active = 1
       AND service_code = ?
       AND (exam_type = ? OR exam_type IN ('A','ANY','ALL','GENERAL'))
       AND (sex = ? OR sex IN ('A','ANY'))
       AND min_age <= ?
       AND max_age >= ?
       AND effective_from <= ?
       AND (effective_to IS NULL OR effective_to = '' OR effective_to >= ?)
     ORDER BY effective_from DESC, id DESC
     LIMIT 1`
  ).bind(service_code, ex, sx, a, a, d, d).first();

  return row || null;
}

async function loadDiscounts(env, valid_on) {
  const d = toIsoDate(valid_on);
  const rows = await env.DB.prepare(
    `SELECT id, name, priority, type, value_bps, value_cents, currency, stackable, conditions_json
     FROM discounts
     WHERE active = 1
       AND effective_from <= ?
       AND (effective_to IS NULL OR effective_to = '' OR effective_to >= ?)
     ORDER BY priority ASC, id ASC`
  ).bind(d, d).all();

  return (rows.results || []).map(x => ({
    ...x,
    type: x.type === "fixed" ? "fixed" : "percent",
    conditions: safeJsonParse(x.conditions_json) || {},
  }));
}

function checkDiscountConditions(cond, ctx) {
  const reason = [];

  if (cond.exam_type) {
    const ok = String(cond.exam_type).trim().toUpperCase() === String(ctx.exam_type).trim().toUpperCase();
    if (!ok) return { ok: false };
    reason.push("exam_type");
  }
  if (cond.client_ruc) {
    const ok = normalizeRuc(cond.client_ruc) === normalizeRuc(ctx.client_ruc);
    if (!ok) return { ok: false };
    reason.push("cliente");
  }
  if (cond.min_subtotal) {
    const minC = moneyToCents(cond.min_subtotal);
    if (ctx.subtotal_cents < minC) return { ok: false };
    reason.push(`subtotal>=${minC}`);
  }
  if (cond.min_qty) {
    const n = clampInt(cond.min_qty, 0, 0, 1_000_000);
    if (ctx.totalQty < n) return { ok: false };
    reason.push(`qty>=${n}`);
  }
  if (cond.min_people) {
    const n = clampInt(cond.min_people, 0, 0, 1_000_000);
    if ((ctx.peopleCount || 0) < n) return { ok: false };
    reason.push(`personas>=${n}`);
  }

  // Apply discount only on subset of services
  let base = ctx.subtotal_cents;
  if (Array.isArray(cond.service_codes) && cond.service_codes.length) {
    const set = new Set(cond.service_codes.map(normCode).filter(Boolean));
    base = ctx.lines
      .filter(x => set.has(normCode(x.service_code)))
      .reduce((s, x) => s + (x.line_total_cents || 0), 0);
    if (base <= 0) return { ok: false };
    reason.push("subset servicios");
  }

  return { ok: true, base_cents: base, reason: reason.length ? `Condiciones: ${reason.join(", ")}` : "Condiciones cumplidas" };
}

async function loadMatrixRules(env, exam_type) {
  const ex = normText(exam_type) || "GENERAL";
  const rows = await env.DB.prepare(
    `SELECT id, name, exam_type, sex, min_age, max_age, services_json, active
     FROM matrix_rules
     WHERE active = 1 AND (exam_type = ? OR exam_type IN ('A','ANY','ALL','GENERAL'))
     ORDER BY id DESC`
  ).bind(ex).all();

  return (rows.results || []).map(r => ({
    id: r.id,
    name: r.name,
    exam_type: r.exam_type,
    sex: normSex(r.sex),
    min_age: clampInt(r.min_age, 0, 0, 120),
    max_age: clampInt(r.max_age, 200, 0, 200),
    services: (safeJsonParse(r.services_json) || []).map(normCode).filter(Boolean),
  }));
}

function pickServicesForEmployee(rules, { sex, age }) {
  const sx = normSex(sex);
  const a = clampInt(age, 0, 0, 200);
  const out = new Set();

  for (const r of rules || []) {
    const sexOk = r.sex === "A" || r.sex === "ANY" || r.sex === sx;
    const ageOk = r.min_age <= a && r.max_age >= a;
    if (!sexOk || !ageOk) continue;

    for (const s of r.services || []) out.add(s);
  }
  return [...out.values()];
}

// ======================
// Bulk validation
// ======================
function validateBulkRows(rows, asOfDate) {
  const asOf = toIsoDate(asOfDate);
  const errors = [];
  const valid_rows = [];

  let i = 0;
  for (const r of rows || []) {
    i++;
    const msgs = [];

    const full_name = String(r?.full_name || r?.name || "").trim();
    const id_number = String(r?.id_number || r?.id || r?.cedula || "").trim();
    const dobRaw = String(r?.dob || r?.fecha_nacimiento || "").trim();
    const sexRaw = String(r?.sex || r?.sexo || "").trim();
    const exam_type = normText(r?.exam_type || r?.tipo_examen) || "GENERAL";
    const email = String(r?.email || "").trim() || null;

    if (!full_name) msgs.push("full_name requerido");
    if (!id_number) msgs.push("id_number requerido");
    const dob = dobRaw ? toIsoDate(dobRaw, true) : null;
    if (!dob) msgs.push("dob inválido (usa YYYY-MM-DD)");

    const sex = normSex(sexRaw);
    if (!sex || (sex !== "M" && sex !== "F")) msgs.push("sex debe ser M o F");

    let age = null;
    if (dob) {
      age = computeAgeYears(dob, asOf);
      if (age < 0 || age > 120) msgs.push("edad fuera de rango");
    }

    if (msgs.length) {
      errors.push({ row: i, messages: msgs });
      continue;
    }

    valid_rows.push({
      full_name,
      id_number,
      dob,
      sex,
      age,
      exam_type,
      email,
    });
  }

  return { valid_rows, errors, summary: { total: rows.length, valid: valid_rows.length, invalid: errors.length } };
}

// ======================
// Discount & matrix normalizers
// ======================
function normalizeDiscount(input, isUpdate = false) {
  const name = String(input?.name || "").trim();
  const type = String(input?.type || "percent").trim().toLowerCase();
  const currency = normText(input?.currency) || "USD";
  const stackable = !!(input?.stackable === true || input?.stackable === 1 || String(input?.stackable) === "1");
  const active = input?.active == null ? 1 : (input.active ? 1 : 0);
  const priority = clampInt(input?.priority, 100, 0, 10_000);

  const effective_from = toIsoDate(String(input?.effective_from || new Date().toISOString().slice(0,10)));
  const effective_to = input?.effective_to ? toIsoDate(String(input?.effective_to)) : null;

  if (!name) return { ok: false, error: "BAD_NAME", message: "name es requerido." };

  const valueRaw = input?.value;
  let value_bps = 0;
  let value_cents = 0;

  if (type === "fixed") {
    value_cents = moneyToCents(valueRaw);
    if (!Number.isFinite(value_cents) || value_cents <= 0) return { ok: false, error: "BAD_VALUE", message: "value (fixed) debe ser > 0." };
  } else {
    // percent
    const p = Number(String(valueRaw ?? "").replace(",", "."));
    if (!Number.isFinite(p) || p <= 0 || p > 100) return { ok: false, error: "BAD_VALUE", message: "value (percent) debe ser 0-100." };
    value_bps = Math.round(p * 100); // 10.5% => 1050 bps
  }

  const conditions = normalizeDiscountConditions(input?.conditions || input?.conditions_json || {});
  return {
    ok: true,
    value: {
      name,
      priority,
      type: type === "fixed" ? "fixed" : "percent",
      value_bps,
      value_cents,
      currency,
      stackable: stackable ? 1 : 0,
      conditions,
      active,
      effective_from,
      effective_to,
    },
  };
}

function normalizeDiscountConditions(c) {
  const obj = typeof c === "string" ? safeJsonParse(c) : c;
  const cond = (obj && typeof obj === "object") ? obj : {};

  const out = {};
  if (cond.exam_type) out.exam_type = normText(cond.exam_type);
  if (cond.client_ruc) out.client_ruc = String(cond.client_ruc).trim();
  if (cond.min_subtotal != null && String(cond.min_subtotal).trim() !== "") out.min_subtotal = String(cond.min_subtotal).trim();
  if (cond.min_qty != null && String(cond.min_qty).trim() !== "") out.min_qty = clampInt(cond.min_qty, 0, 0, 1_000_000);
  if (cond.min_people != null && String(cond.min_people).trim() !== "") out.min_people = clampInt(cond.min_people, 0, 0, 1_000_000);

  // service_codes: accept array or comma-separated string
  if (cond.service_codes) {
    if (Array.isArray(cond.service_codes)) {
      out.service_codes = cond.service_codes.map(normCode).filter(Boolean);
    } else {
      out.service_codes = String(cond.service_codes).split(",").map(s => normCode(s)).filter(Boolean);
    }
  }
  return out;
}

function normalizeMatrixRule(input) {
  const name = String(input?.name || "").trim() || "Regla";
  const exam_type = normText(input?.exam_type) || "GENERAL";
  const sex = normSex(input?.sex || "A");
  const min_age = clampInt(input?.min_age, 0, 0, 120);
  const max_age = clampInt(input?.max_age, 200, 0, 200);
  const notes = String(input?.notes || "").trim() || null;
  const active = input?.active == null ? 1 : (input.active ? 1 : 0);

  let services = input?.services;
  if (!services && input?.services_json) services = input.services_json;
  if (typeof services === "string") {
    // allow comma-separated
    services = services.split(",").map(s => normCode(s)).filter(Boolean);
  }
  if (Array.isArray(services)) {
    services = services.map(normCode).filter(Boolean);
  } else {
    services = [];
  }

  if (!services.length) return { ok: false, error: "NO_SERVICES", message: "services es requerido (al menos 1 código)." };

  return { ok: true, value: { name, exam_type, sex, min_age, max_age, services, notes, active } };
}

// ======================
// Audit
// ======================
async function audit(env, { actor, action, entity, entity_id, before, after, rid }) {
  try {
    if (!env?.DB) return;
    await env.DB.prepare(
      `INSERT INTO audit_log (ts, actor, action, entity, entity_id, before_json, after_json, rid)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
    ).bind(
      new Date().toISOString(),
      String(actor || ""),
      String(action || ""),
      String(entity || ""),
      String(entity_id || ""),
      before ? JSON.stringify(before) : null,
      after ? JSON.stringify(after) : null,
      String(rid || "")
    ).run();
  } catch {
    // best effort
  }
}

// ======================
// RBAC + Workflow helpers
// ======================
const SCHEMA_CACHE = {
  checked: false,
  rbac: false,
  approvals: false,
  branding: false,
  crm: false,
  quotes_status: false,
};

async function getSchemaState(env) {
  if (SCHEMA_CACHE.checked) return SCHEMA_CACHE;
  if (!env?.DB) {
    SCHEMA_CACHE.checked = true;
    return SCHEMA_CACHE;
  }

  try {
    const tables = await env.DB.prepare(`SELECT name FROM sqlite_master WHERE type='table'`).all();
    const set = new Set((tables.results || []).map(r => r.name));
    SCHEMA_CACHE.rbac = set.has("users");
    SCHEMA_CACHE.approvals = set.has("quote_approvals");
    SCHEMA_CACHE.branding = set.has("client_branding");
    SCHEMA_CACHE.crm = set.has("crm_sync_log");
  } catch {
    // ignore
  }

  try {
    const info = await env.DB.prepare(`PRAGMA table_info(quotes)`).all();
    const cols = new Set((info.results || []).map(r => r.name));
    SCHEMA_CACHE.quotes_status = cols.has("status");
  } catch {
    // ignore
  }

  SCHEMA_CACHE.checked = true;
  return SCHEMA_CACHE;
}

function normRole(role) {
  const r = String(role || "").trim().toLowerCase();
  if (!r) return null;
  if (["admin", "administrator"].includes(r)) return "admin";
  if (["supervisor", "manager", "lead"].includes(r)) return "supervisor";
  if (["advisor", "asesor", "sales", "seller", "user"].includes(r)) return "advisor";
  return null;
}

async function getAuthzContext(env, claims, schema) {
  const oid = String(claims?.oid || "").trim();
  if (!oid) throw authErr("Missing oid claim", 401);

  const entraAdmin = isAdmin(claims, env);
  const entraFallback = entraAdmin && String(env.ENABLE_ENTRA_ADMIN_FALLBACK ?? "1") !== "0";

  const email = String(claims?.preferred_username || claims?.upn || "").trim() || null;
  const display_name = String(claims?.name || "").trim() || null;

  let dbUser = null;
  let dbRole = null;
  let supervisor_oid = null;

  if (schema?.rbac) {
    dbUser = await env.DB.prepare(
      `SELECT oid, email, display_name, role, supervisor_oid, active
       FROM users
       WHERE oid = ?`
    ).bind(oid).first();

    if (!dbUser) {
      const now = new Date().toISOString();
      const bootstrapRole = entraFallback && String(env.BOOTSTRAP_ADMIN_ON_FIRST_LOGIN ?? "1") !== "0" ? "admin" : "advisor";
      await env.DB.prepare(
        `INSERT INTO users (oid, email, display_name, role, supervisor_oid, active, created_at, created_by)
         VALUES (?, ?, ?, ?, NULL, 1, ?, ?)`
      ).bind(oid, email, display_name, bootstrapRole, now, oid).run();

      dbUser = { oid, email, display_name, role: bootstrapRole, supervisor_oid: null, active: 1 };
    } else {
      // keep profile fresh
      const needsEmail = (!dbUser.email || dbUser.email === null) && email;
      const needsName = (!dbUser.display_name || dbUser.display_name === null) && display_name;
      if (needsEmail || needsName) {
        const now = new Date().toISOString();
        await env.DB.prepare(
          `UPDATE users SET email=?, display_name=?, updated_at=?, updated_by=? WHERE oid=?`
        ).bind(
          needsEmail ? email : dbUser.email,
          needsName ? display_name : dbUser.display_name,
          now,
          oid,
          oid
        ).run();
        dbUser.email = needsEmail ? email : dbUser.email;
        dbUser.display_name = needsName ? display_name : dbUser.display_name;
      }
    }

    if (Number(dbUser.active) === 0) {
      const e = new Error("USER_INACTIVE");
      e.status = 403;
      throw e;
    }

    dbRole = normRole(dbUser.role) || "advisor";
    supervisor_oid = dbUser.supervisor_oid || null;
  } else {
    dbRole = entraFallback ? "admin" : "advisor";
    dbUser = { oid, email, display_name, role: dbRole, supervisor_oid: null, active: 1 };
  }

  const isAdminEffective = dbRole === "admin" || entraFallback;
  const isSupervisorEffective = !isAdminEffective && dbRole === "supervisor";
  const role = isAdminEffective ? "admin" : (isSupervisorEffective ? "supervisor" : "advisor");

  const maxAdvisor = clampInt(env.ADVISOR_MAX_DISCOUNT_BPS, 500, 0, 10_000);
  const maxSupervisor = clampInt(env.SUPERVISOR_MAX_DISCOUNT_BPS, 1500, 0, 10_000);
  const max_manual_discount_bps = role === "advisor" ? maxAdvisor : (role === "supervisor" ? maxSupervisor : null);

  return {
    rbac_enabled: !!schema?.rbac,
    user: {
      oid,
      email: dbUser.email || email,
      display_name: dbUser.display_name || display_name,
      role,
      supervisor_oid,
      active: 1,
    },
    role,
    isAdmin: isAdminEffective,
    isSupervisor: role === "supervisor",
    isAdvisor: role === "advisor",
    max_manual_discount_bps,
    entra_admin: entraAdmin,
  };
}

function assertAdminRole(authz) {
  if (!authz?.isAdmin) {
    const e = new Error("ADMIN_REQUIRED");
    e.status = 403;
    throw e;
  }
}

function assertSupervisorOrAdmin(authz) {
  if (!authz?.isAdmin && !authz?.isSupervisor) {
    const e = new Error("SUPERVISOR_REQUIRED");
    e.status = 403;
    throw e;
  }
}

async function assertCanAccessQuote(env, schema, authz, quoteRow) {
  if (authz?.isAdmin) return true;
  const created_by = String(quoteRow?.created_by || "");
  if (created_by && authz?.user?.oid && created_by === authz.user.oid) return true;

  if (authz?.isSupervisor && schema?.rbac && created_by) {
    const ok = await env.DB.prepare(
      `SELECT 1 FROM users WHERE oid = ? AND supervisor_oid = ? AND active = 1 LIMIT 1`
    ).bind(created_by, authz.user.oid).first();
    if (ok) return true;
  }

  const e = new Error("FORBIDDEN");
  e.status = 403;
  throw e;
}

async function setQuoteStatus(env, schema, quote_id, { status, updated_at, updated_by, approved_at, approved_by }) {
  if (!schema?.quotes_status) return false;

  const row = await env.DB.prepare(
    `SELECT id, status, approved_at, approved_by, updated_at, updated_by, result_json
     FROM quotes
     WHERE id = ?`
  ).bind(quote_id).first();
  if (!row) return false;

  const newStatus = String(status || row.status || "draft");
  const newUpdatedAt = updated_at != null ? updated_at : (new Date().toISOString());
  const newUpdatedBy = updated_by != null ? updated_by : row.updated_by;
  const newApprovedAt = approved_at !== undefined ? approved_at : row.approved_at;
  const newApprovedBy = approved_by !== undefined ? approved_by : row.approved_by;

  let result = safeJsonParse(row.result_json);
  if (result && typeof result === "object") {
    result.status = newStatus;
    if (result.workflow && typeof result.workflow === "object") {
      // keep original values; only flip required flag when approved/rejected
      if (newStatus === "approved") result.workflow.approval_required = false;
      if (newStatus === "draft" && String(row.status) === "pending_approval") {
        // likely rejected
        result.workflow.approval_required = false;
      }
    }
    if (result.meta && typeof result.meta === "object") {
      if (newApprovedAt) result.meta.approved_at = newApprovedAt;
      if (newApprovedBy) result.meta.approved_by = newApprovedBy;
    }
  }

  await env.DB.prepare(
    `UPDATE quotes
       SET status=?, updated_at=?, updated_by=?, approved_at=?, approved_by=?, result_json=?
     WHERE id=?`
  ).bind(
    newStatus,
    newUpdatedAt,
    newUpdatedBy,
    newApprovedAt,
    newApprovedBy,
    result ? JSON.stringify(result) : row.result_json,
    quote_id
  ).run();

  return true;
}

// ======================
// Auth (Entra ID)
// ======================
function authErr(message, status = 401) {
  const e = new Error(message);
  e.status = status;
  return e;
}

const JWKS_CACHE = new Map();

async function verifyEntraJwt(req, env, ctx) {
  const auth = req.headers.get("Authorization") || "";
  const m = auth.match(/^Bearer\s+(.+)$/i);
  if (!m) throw authErr("Missing bearer token", 401);

  const token = m[1].trim();
  const parts = token.split(".");
  if (parts.length !== 3) throw authErr("Invalid token format", 401);
  const [h64, p64, s64] = parts;

  const header = JSON.parse(b64urlToText(h64));
  const payload = JSON.parse(b64urlToText(p64));

  const kid = header?.kid;
  const alg = header?.alg;
  if (alg !== "RS256") throw authErr(`Unsupported alg: ${alg || "none"}`, 401);
  if (!kid) throw authErr("Missing kid", 401);

  const tenant = (env.ENTRA_TENANT_ID || "").trim();
  if (!tenant) throw authErr("Missing ENTRA_TENANT_ID", 500);

  const now = Math.floor(Date.now() / 1000);
  if (payload?.tid && payload.tid !== tenant) throw authErr("tid mismatch", 401);
  if (!payload?.exp || payload.exp < now) throw authErr("token expired", 401);
  if (payload?.nbf && payload.nbf > now + 60) throw authErr("token not active yet", 401);

  const audExpected = (env.ENTRA_API_AUDIENCE || "").trim();
  if (!audExpected) throw authErr("Missing ENTRA_API_AUDIENCE", 500);
  if (!audMatch(payload?.aud, audExpected)) throw authErr("aud mismatch", 401);

  const iss = String(payload?.iss || "");
  const issOk =
    iss === `https://login.microsoftonline.com/${tenant}/v2.0` ||
    iss === `https://sts.windows.net/${tenant}/`;
  if (!issOk) throw authErr("iss mismatch", 401);

  const required = (env.ENTRA_REQUIRED_SCOPE || "").trim();
  if (!required) throw authErr("Missing ENTRA_REQUIRED_SCOPE", 500);
  if (!hasScopeOrRole(payload, required)) throw authErr("Missing required scope/role", 403);

  const key = await getSigningKeyForKid(tenant, kid, env, ctx);
  const data = new TextEncoder().encode(`${h64}.${p64}`);
  const sig = b64urlToBytes(s64);

  const ok = await crypto.subtle.verify(
    { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" },
    key,
    sig,
    data
  );
  if (!ok) throw authErr("Invalid signature", 401);

  return payload;
}

function isAdmin(claims, env) {
  const adminRole = (env.ENTRA_ADMIN_ROLE || "").trim();
  if (!adminRole) return false;
  return hasScopeOrRole(claims, adminRole);
}

function assertAdmin(claims, env) {
  if (!isAdmin(claims, env)) {
    const e = new Error("ADMIN_REQUIRED");
    e.status = 403;
    throw e;
  }
}

function hasScopeOrRole(payload, required) {
  const scp = String(payload?.scp || "");
  const roles = Array.isArray(payload?.roles) ? payload.roles : [];
  const scopes = scp.split(" ").map(s => s.trim()).filter(Boolean);
  return scopes.includes(required) || roles.includes(required);
}

function audMatch(audClaim, expected) {
  if (!audClaim) return false;
  const exp = String(expected || "").trim();
  if (!exp) return false;

  const candidates = new Set([exp]);
  // If expected is GUID, accept api://GUID
  if (/^[0-9a-fA-F-]{36}$/.test(exp)) candidates.add(`api://${exp}`);

  const matchOne = (v) => {
    const s = String(v || "").trim();
    if (!s) return false;
    if (candidates.has(s)) return true;
    // If expected is api://GUID, accept GUID
    if (s.startsWith("api://") && candidates.has(s.replace("api://", ""))) return true;
    return false;
  };

  if (typeof audClaim === "string") return matchOne(audClaim);
  if (Array.isArray(audClaim)) return audClaim.some(matchOne);
  return false;
}

async function getSigningKeyForKid(tenant, kid, env, ctx) {
  const ttlS = num(env.CACHE_TTL_JWKS_S, 3600);
  const cached = JWKS_CACHE.get(tenant);
  const nowMs = Date.now();

  if (cached && cached.expMs > nowMs && cached.keyByKid?.has(kid)) {
    return cached.keyByKid.get(kid);
  }

  const jwksUrl = `https://login.microsoftonline.com/${tenant}/discovery/v2.0/keys`;
  const cacheKey = new Request(jwksUrl, { method: "GET" });

  let jwksText = null;
  try {
    const hit = await caches.default.match(cacheKey);
    if (hit) jwksText = await hit.text();
  } catch {
    // ignore
  }

  if (!jwksText) {
    const r = await fetch(jwksUrl, { headers: { accept: "application/json" } });
    jwksText = await r.text();
    if (!r.ok) throw authErr(`JWKS fetch failed: ${r.status}`, 401);

    try {
      ctx?.waitUntil?.(
        caches.default.put(
          cacheKey,
          new Response(jwksText, {
            headers: { "content-type": "application/json", "Cache-Control": `public, max-age=${ttlS}` },
          })
        )
      );
    } catch {
      // ignore
    }
  }

  const jwks = JSON.parse(jwksText);
  const keys = Array.isArray(jwks?.keys) ? jwks.keys : [];
  const keyByKid = new Map();

  for (const k of keys) {
    if (!k?.kid) continue;
    if (k.kty !== "RSA" || !k.n || !k.e) continue;

    const cryptoKey = await crypto.subtle.importKey(
      "jwk",
      k,
      { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" },
      false,
      ["verify"]
    );
    keyByKid.set(k.kid, cryptoKey);
  }

  JWKS_CACHE.set(tenant, { expMs: nowMs + ttlS * 1000, keyByKid });

  const key = keyByKid.get(kid);
  if (!key) throw authErr("kid not found in JWKS", 401);
  return key;
}

// ======================
// CORS
// ======================
function corsHeaders(req, env) {
  const origin = req.headers.get("Origin") || "";
  const allow = String(env.ALLOWED_ORIGINS || "")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);

  const h = {
    "Access-Control-Allow-Methods": "GET,POST,PUT,PATCH,DELETE,OPTIONS",
    "Access-Control-Allow-Headers": "content-type,authorization",
    "Access-Control-Max-Age": "86400",
    Vary: "Origin",
  };

  if (!allow.length) {
    if (origin) h["Access-Control-Allow-Origin"] = origin;
  } else {
    if (allow.includes(origin)) h["Access-Control-Allow-Origin"] = origin;
  }
  return h;
}

// ======================
// Responses + parsing
// ======================
function j(obj, status = 200, headers = {}) {
  return new Response(JSON.stringify(obj, null, 2), {
    status,
    headers: { "content-type": "application/json; charset=utf-8", ...headers },
  });
}

async function readJson(req, maxBytes) {
  const txt = await readBodyLimited(req, maxBytes);
  let body = null;
  try {
    body = JSON.parse(txt || "{}");
  } catch {
    const e = new Error("BAD_JSON");
    e.status = 400;
    throw e;
  }
  return body;
}

async function readBodyLimited(req, maxBytes) {
  const cl = req.headers.get("content-length");
  if (cl && Number(cl) > maxBytes) {
    const e = new Error(`BODY_TOO_LARGE ${cl} > ${maxBytes}`);
    e.status = 413;
    throw e;
  }

  const reader = req.body?.getReader?.();
  if (!reader) return await req.text();

  const chunks = [];
  let total = 0;
  while (true) {
    const { value, done } = await reader.read();
    if (done) break;
    total += value.byteLength;
    if (total > maxBytes) {
      const e = new Error(`BODY_TOO_LARGE > ${maxBytes}`);
      e.status = 413;
      throw e;
    }
    chunks.push(value);
  }

  const merged = new Uint8Array(total);
  let off = 0;
  for (const c of chunks) {
    merged.set(c, off);
    off += c.byteLength;
  }
  return new TextDecoder().decode(merged);
}

// ======================
// Helpers
// ======================
function safeUserFromClaims(claims) {
  return {
    oid: claims?.oid || null,
    upn: claims?.preferred_username || claims?.upn || null,
    name: claims?.name || null,
    tid: claims?.tid || null,
  };
}

function actorId(claims) {
  return String(claims?.oid || claims?.preferred_username || claims?.upn || "unknown");
}

function mkRid() {
  return Math.random().toString(16).slice(2) + "-" + Date.now().toString(16);
}

function cryptoRandomId() {
  const a = new Uint8Array(16);
  crypto.getRandomValues(a);
  return [...a].map((b) => b.toString(16).padStart(2, "0")).join("");
}

function num(v, def) {
  const n = Number(v);
  return Number.isFinite(n) ? n : def;
}

function clampInt(v, def, min, max) {
  const n = Number(v);
  const x = Number.isFinite(n) ? Math.floor(n) : def;
  return Math.max(min, Math.min(max, x));
}

function normText(s) {
  const t = String(s || "").trim();
  return t ? t.toUpperCase() : "";
}

function normCode(s) {
  return String(s || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "_")
    .replace(/[^\w\-_.]/g, "");
}

function normSex(s) {
  const t = String(s || "").trim().toUpperCase();
  if (!t) return "A";
  if (t === "M" || t === "F") return t;
  if (t === "A" || t === "ANY" || t === "ALL") return "A";
  if (t.startsWith("MASC")) return "M";
  if (t.startsWith("FEM")) return "F";
  return "A";
}

function normalizeRuc(s) {
  return String(s || "").replace(/\D+/g, "").trim();
}

function toIsoDate(s, allowSlash = false) {
  const raw = String(s || "").trim();
  if (!raw) return null;

  // Accept YYYY-MM-DD
  let m = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return raw;

  // Accept YYYY/MM/DD (optional)
  if (allowSlash) {
    m = raw.match(/^(\d{4})\/(\d{2})\/(\d{2})$/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  }

  // Accept DD/MM/YYYY
  m = raw.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) return `${m[3]}-${m[2]}-${m[1]}`;

  // Fallback Date parse
  const d = new Date(raw);
  if (!Number.isFinite(d.getTime())) return null;
  return d.toISOString().slice(0, 10);
}

function computeAgeYears(dobIso, asOfIso) {
  const dob = new Date(dobIso + "T00:00:00Z");
  const asOf = new Date(asOfIso + "T00:00:00Z");
  let age = asOf.getUTCFullYear() - dob.getUTCFullYear();
  const m = asOf.getUTCMonth() - dob.getUTCMonth();
  if (m < 0 || (m === 0 && asOf.getUTCDate() < dob.getUTCDate())) age--;
  return age;
}

function moneyToCents(v) {
  if (v == null) return NaN;
  if (typeof v === "number") return Math.round(v * 100);

  const s = String(v).trim();
  if (!s) return NaN;

  // Remove currency symbols and spaces
  const cleaned = s.replace(/[^\d,.\-]/g, "").replace(/\s+/g, "");

  // If both comma and dot exist, assume comma is thousands separator
  const hasComma = cleaned.includes(",");
  const hasDot = cleaned.includes(".");
  let norm = cleaned;

  if (hasComma && hasDot) {
    norm = norm.replace(/,/g, "");
  } else if (hasComma && !hasDot) {
    // comma as decimal separator
    norm = norm.replace(",", ".");
  }

  const n = Number(norm);
  if (!Number.isFinite(n)) return NaN;
  return Math.round(n * 100);
}

function safeJsonParse(s) {
  if (s == null) return null;
  try {
    return JSON.parse(String(s));
  } catch {
    return null;
  }
}

function b64urlToText(b64url) {
  const bytes = b64urlToBytes(b64url);
  return new TextDecoder().decode(bytes);
}

function b64urlToBytes(b64url) {
  const b64 = b64url.replace(/-/g, "+").replace(/_/g, "/") + "===".slice((b64url.length + 3) % 4);
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

async function sha256Hex(text) {
  const bytes = new TextEncoder().encode(String(text || ""));
  const hash = await crypto.subtle.digest("SHA-256", bytes);
  const arr = new Uint8Array(hash);
  return [...arr].map(b => b.toString(16).padStart(2, "0")).join("");
}

// ======================
// Rate limit (best-effort)
// ======================
const RATE = new Map();
function rateLimit(req, limit = 30, windowMs = 60_000) {
  const ip =
    req.headers.get("CF-Connecting-IP") ||
    req.headers.get("x-forwarded-for")?.split(",")[0]?.trim() ||
    "unknown";

  const now = Date.now();
  const rec = RATE.get(ip) || { ts: now, n: 0 };
  if (now - rec.ts > windowMs) {
    rec.ts = now;
    rec.n = 0;
  }
  rec.n++;
  RATE.set(ip, rec);

  if (rec.n > limit) {
    const retryAfterMs = windowMs - (now - rec.ts);
    return { ok: false, retryAfterSec: Math.max(1, Math.ceil(retryAfterMs / 1000)) };
  }
  return { ok: true, retryAfterSec: 0 };
}
