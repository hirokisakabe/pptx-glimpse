export const DEV_EDITOR_HTTP_SCRIPT = `    function postJson(url, payload) {
      return fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload || {})
      }).then(function (res) {
        return res.json().then(function (data) {
          if (!res.ok) throw new Error(data.error || "request failed");
          return data;
        });
      });
    }

    function escapeHtmlClient(value) {
      return String(value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
    }

`;
