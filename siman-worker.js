// סימן — Worker שמגיש את האתר מכתובת Cloudflare, בלי לחשוף את GitHub
// כתובת: https://siman.siman-app.workers.dev
const ORIGIN = "https://lial-apps.github.io/siman";

const NOT_FOUND = `<!doctype html><html lang="he" dir="rtl"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>סימן — הדף לא נמצא</title>
<style>body{margin:0;min-height:100vh;display:grid;place-items:center;background:#eef1f6;color:#16213b;
font-family:Rubik,Heebo,Arial,sans-serif;text-align:center;padding:24px}
h1{font-size:56px;margin:0 0 8px}h1 span{background:linear-gradient(transparent 55%,#ffd93d 55%);padding:0 6px}
p{color:#5d6782;font-size:18px}a{display:inline-block;margin-top:14px;background:#ffd93d;color:#16213b;
text-decoration:none;font-weight:700;padding:14px 28px;border-radius:14px}</style></head>
<body><div><h1><span>סימן</span></h1><p>הדף שחיפשתם לא נמצא.</p><a href="/">לפתיחת סימן</a></div></body></html>`;

export default {
  async fetch(request) {
    const url = new URL(request.url);
    if (request.method !== "GET" && request.method !== "HEAD") {
      return new Response("Method not allowed", { status: 405 });
    }

    // אם ה-manifest או קישור כלשהו מכילים /siman/ — מורידים את הקידומת
    let path = url.pathname.replace(/^\/siman(?=\/|$)/, "") || "/";

    // מעבירים Range כדי שהסרטון יתנגן באייפון (Safari דורש הזרמה בחלקים)
    const fwd = new Headers();
    for (const h of ["range", "if-none-match", "if-modified-since", "accept", "accept-encoding"]) {
      const v = request.headers.get(h); if (v) fwd.set(h, v);
    }

    const res = await fetch(ORIGIN + path + url.search, {
      method: request.method,
      headers: fwd,
      redirect: "manual",
      cf: { cacheEverything: true, cacheTtlByStatus: { "200-299": 300, "300-599": 0 } }, // שגיאות לא נשמרות במטמון
    });

    // דף 404 של סימן במקום הדף של GitHub
    if (res.status === 404) {
      return new Response(NOT_FOUND, {
        status: 404,
        headers: { "Content-Type": "text/html; charset=utf-8", "Cache-Control": "no-store" },
      });
    }

    const headers = new Headers(res.headers);

    // הפניות של GitHub — משכתבים לכתובת שלנו
    const loc = headers.get("Location");
    if (loc) {
      headers.set("Location", loc.replace(/^https?:\/\/lial-apps\.github\.io\/siman/, url.origin));
    }

    // לא חושפים כותרות של GitHub
    ["x-github-request-id", "x-served-by", "x-fastly-request-id", "via", "x-proxy-cache"].forEach(h => headers.delete(h));

    // sw.js ו-manifest תמיד טריים, כדי שעדכונים יגיעו מהר
    if (/\/(sw\.js|manifest\.webmanifest|index\.html|install\.html)?$/.test(path)) {
      headers.set("Cache-Control", "no-cache");
    }

    return new Response(res.body, { status: res.status, statusText: res.statusText, headers });
  },
};
