export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === "/api" || url.pathname.startsWith("/api/")) {
      const email = request.headers.get(
        "cf-access-authenticated-user-email"
      );

      if (!email) {
        return jsonResponse(
          {
            ok: false,
            error: "Cloudflare Access identity not found"
          },
          403
        );
      }

      // TEMPORARY DEBUG ENDPOINT
      if (url.pathname === "/api/debug") {
        return jsonResponse({
          ok: true,
          authenticated: true,
          user: email
        });
      }

      if (!env.TASKBOARD_API_SECRET) {
        return jsonResponse(
          {
            ok: false,
            error: "Worker secret is not configured"
          },
          500
        );
      }

      if (!env.APPS_SCRIPT_URL) {
        return jsonResponse(
          {
            ok: false,
            error: "Apps Script URL is not configured"
          },
          500
        );
      }

      try {
        let body = {};

        if (request.method === "GET") {
          body = Object.fromEntries(url.searchParams.entries());
        } else if (request.method === "POST") {
          const contentType =
            request.headers.get("content-type") || "";

          if (contentType.includes("application/json")) {
            body = await request.json();
          } else {
            body = {};
          }
        } else {
          return jsonResponse(
            {
              ok: false,
              error: "Method not allowed"
            },
            405
          );
        }

        body.user_email = email;
        body.secret = env.TASKBOARD_API_SECRET;

        const response = await fetch(env.APPS_SCRIPT_URL, {
          method: "POST",
          headers: {
            "Content-Type": "application/json"
          },
          body: JSON.stringify(body)
        });

        const responseText = await response.text();

        return new Response(responseText, {
          status: response.status,
          headers: {
            "Content-Type":
              response.headers.get("Content-Type") ||
              "application/json"
          }
        });

      } catch (error) {
        return jsonResponse(
          {
            ok: false,
            error: "Failed to contact Apps Script"
          },
          502
        );
      }
    }

    return env.ASSETS.fetch(request);
  }
};

function jsonResponse(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      "Content-Type": "application/json"
    }
  });
}
