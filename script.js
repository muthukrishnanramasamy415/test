let excelRows = [];
let zbxToken = "";

// DARK MODE
function toggleDark() { document.body.classList.toggle("dark"); }

// LOGIN
async function testLogin() {
    const payload = {
        jsonrpc: "2.0",
        method: "user.login",
        params: { username: zbxUser.value, password: zbxPass.value },
        id: 1
    };
    const res = await fetch(zbxUrl.value, {
        method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify(payload)
    });
    const json = await res.json();
    loginStatus.innerHTML = json.result ?
        "<span class='success'>Login Successful!</span>" :
        "<span class='error'>Login Failed</span>";
    if(json.result) zbxToken = json.result;
}

// READ EXCEL
fileInput.addEventListener("change", e => {
    const reader = new FileReader();
    reader.onload = ev => {
        const wb = XLSX.read(new Uint8Array(ev.target.result), {type: "array"});
        excelRows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        previewExcel();
    };
    reader.readAsArrayBuffer(e.target.files[0]);
});

// PREVIEW
function previewExcel() {
    previewTable.innerHTML = "";
    if(excelRows.length === 0) return;

    let header = "<tr>";
    Object.keys(excelRows[0]).forEach(k => header += `<th>${k}</th>`);
    header += "</tr>";
    previewTable.innerHTML = header;

    excelRows.slice(0, 10).forEach(r => {
        let row = "<tr>";
        Object.values(r).forEach(v => row += `<td>${v}</td>`);
        row += "</tr>";
        previewTable.innerHTML += row;
    });
}

// START IMPORT
async function startImport() {
    if(!zbxToken) return alert("Login first.");
    if(excelRows.length === 0) return alert("Upload Excel file.");

    progressContainer.style.display = "block";
    resultTable.innerHTML = `
        <tr>
            <th>Hostname</th>
            <th>IP</th>
            <th>Agent</th>
            <th>Status</th>
        </tr>`;

    for(let i=0;i<excelRows.length;i++){
        const row = excelRows[i];
        const resp = await createHost(row);

        const status = resp.result ?
            "<span class='success'>Success</span>" :
            `<span class='error'>${JSON.stringify(resp.error)}</span>`;

        resultTable.innerHTML += `
            <tr>
                <td>${row.Hostname}</td>
                <td>${row.IPAddress}</td>
                <td>${row.AgentType}</td>
                <td>${status}</td>
            </tr>`;

        let pct = Math.round(((i+1)/excelRows.length)*100);
        progressBar.style.width = pct+"%";
        progressBar.innerHTML = pct+"%";
        await new Promise(r=>setTimeout(r,5));
    }

    alert("Import Completed");
}

// CREATE HOST (FIXED)
async function createHost(row) {
    const url = zbxUrl.value;

    const groups = String(row.GroupIDs || "")
        .split(",").map(x=>x.trim()).filter(x=>x)
        .map(x => ({ groupid: x }));

    const templates = String(row.TemplateIDs || "")
        .split(",").map(x=>x.trim()).filter(x=>x)
        .map(x => ({ templateid: x }));

    let interfaces = [];
    switch(String(row.AgentType).toUpperCase()){

        case "AGENT":
            interfaces.push({
                type: 1, main: 1, useip: 1,
                ip: row.IPAddress, dns: "", port: "10050"
            });
            break;

        case "SNMP":
            interfaces.push({
                type: 2, main: 1, useip: 1,
                ip: row.IPAddress, dns: "", port: "161",
                details: { version: 2, bulk: 1, community: row.SNMPString, max_repetitions: 10 }
            });
            break;

        case "ACTIVE":
            interfaces = [];
            break;
    }

    const payload = {
        jsonrpc: "2.0",
        method: "host.create",
        params: {
            host: row.Hostname,
            name: row.Hostname,
            groups: groups,
            templates: templates,
            interfaces: interfaces,
            proxy_hostid: Number(row.ProxyID)
        },
        auth: zbxToken,
        id: Math.floor(Math.random()*9999)
    };

    const res = await fetch(url, {
        method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify(payload)
    });

    return await res.json();
}

// EXPORT
function exportFinal(){
    const wb = XLSX.utils.table_to_book(resultTable, {sheet: "Results"});
    XLSX.writeFile(wb, "zabbix_results.xlsx");
}
