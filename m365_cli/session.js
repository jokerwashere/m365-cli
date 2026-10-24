async ({op, headers: captured, account}) => {
    const hosts = ["outlook.office.com", "outlook.office365.com", "outlook.cloud.microsoft", "outlook.live.com"];
    if (location.protocol !== "https:" || !hosts.includes(location.hostname) || (location.port && location.port !== "443")) {
        return {authenticated: false, error: "login_required"};
    }
    // Auth stays in this dedicated Chrome profile or process memory. Never return
    // bearer/canary values in CLI output and never write a separate token cache.
    const saved = window.__outlookAccessHeaders || {};
    const headers = {...saved, ...captured};
    const decode = token => {
        try { return JSON.parse(atob(token.replace(/^Bearer /i, "").split(".")[1].replace(/-/g, "+").replace(/_/g, "/"))); }
        catch { return {}; }
    };
    let claims = decode(headers.authorization || "");
    if (claims.exp && claims.exp * 1000 < Date.now() + 30000) {
        delete headers.authorization;
        claims = {};
    }
    for (const storage of [localStorage, sessionStorage]) {
        for (const key of Object.keys(storage)) {
            if (!/token/i.test(key)) continue;
            try {
                const item = JSON.parse(storage.getItem(key));
                const token = item.secret || item.token;
                if (typeof token !== "string") continue;
                const c = decode(token);
                const tokenEmail = c.preferred_username || c.upn || c.unique_name || "";
                if (account && tokenEmail && tokenEmail.toLowerCase() !== account.toLowerCase()) continue;
                const audiences = Array.isArray(c.aud) ? c.aud : [c.aud];
                if (audiences.some(a => ["https://outlook.office.com", "https://outlook.office365.com", "00000002-0000-0ff1-ce00-000000000000"].includes(a)) && c.exp * 1000 > Date.now() + 30000) {
                    if (!headers.authorization || (c.exp || 0) > (claims.exp || 0)) {
                        headers.authorization = `Bearer ${token}`;
                        claims = c;
                    }
                }
            } catch {}
        }
    }
    for (const pair of document.cookie.split(";")) {
        const equal = pair.indexOf("=");
        if (["X-OWA-CANARY", "OWA-CANARY", "XOWACANARY"].includes(pair.slice(0, equal).trim().toUpperCase())) {
            headers["x-owa-canary"] = decodeURIComponent(pair.slice(equal + 1));
        }
    }
    let email = claims.preferred_username || claims.upn || claims.unique_name || "";
    const settings = window.owa?.sessionData || window.owaSettings || {};
    email = settings.primarySmtpAddress || email;
    const anchor = headers["x-anchormailbox"] || "";
    if (!email && /^SMTP:/i.test(anchor)) email = anchor.slice(5);
    if (!email.includes("@")) email = "";
    window.__outlookAccessHeaders = headers;
    const authenticated = !!(headers.authorization || headers["x-owa-canary"]);
    if (op === "status") return {authenticated, account: email || null};
    if (!authenticated) return {status: 401, body: {error: "login_required"}};
    if (email && account && email.toLowerCase() !== account.toLowerCase()) return {status: 409, body: {error: "account_mismatch"}};
    // Bootstrap performs exactly one read. All product operations use direct HTTP.
    if (op !== "bootstrap") throw new Error("Unapproved bootstrap operation");
    const base = location.hostname === "outlook.cloud.microsoft" ? "/owa/service.svc" : "/owa/0/service.svc";
    const action = "GetOwaUserConfiguration";
    const requestHeaders = {...headers, Accept: "application/json", "Content-Type": "application/json; charset=utf-8",
        "X-Req-Source": "Mail", "X-OWA-CorrelationId": crypto.randomUUID(), Prefer: 'IdType="ImmutableId"', Action: action};
    if (account || email) requestHeaders["x-anchormailbox"] = `SMTP:${account || email}`;
    const body = {__type:"GetOwaUserConfigurationJsonRequest:#Exchange",
        Header:{__type:"JsonRequestHeaders:#Exchange",RequestServerVersion:"V2018_01_08",
            TimeZoneContext:{__type:"TimeZoneContext:#Exchange",TimeZoneDefinition:{__type:"TimeZoneDefinitionType:#Exchange",Id:"UTC"}}},
        Body:{__type:"GetOwaUserConfigurationRequest:#Exchange"}};
    const response = await fetch(`${base}?action=${action}&app=Mail&n=0`, {method:"POST", headers:requestHeaders,
        credentials:"include", redirect:"error", signal:AbortSignal.timeout(45000), body:JSON.stringify(body)});
    const reader=response.body.getReader(); let size=0; const chunks=[];
    for (;;) {const {done,value}=await reader.read(); if(done)break;
        size+=value.length; if(size>8*1024*1024){await reader.cancel();throw new Error("Bootstrap response limit");}
        chunks.push(value);}
    const bytes=new Uint8Array(size); let at=0;
    for(const chunk of chunks){bytes.set(chunk,at);at+=chunk.length;}
    let result; try {result=JSON.parse(new TextDecoder().decode(bytes));}
    catch {result={error:"non_json_response"};}
    return {status:response.status,body:result,retry_after:response.headers.get("retry-after")};
}
