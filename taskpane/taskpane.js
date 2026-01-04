document.getElementById("loadXml").onclick = async () => {
    const [fileHandle] = await window.showOpenFilePicker({
        types: [{ description: 'XML Files', accept: { 'text/xml': ['.xml'] } }]
    });

    const file = await fileHandle.getFile();
    const text = await file.text();

    // Parse XML
    const parser = new DOMParser();
    const xml = parser.parseFromString(text, "application/xml");

    const hosts = [...xml.querySelectorAll("host")].map(host => {
        const ip = host.querySelector("address")?.getAttribute("addr");
        const ports = [...host.querySelectorAll("port[state][state='open']")].map(p => p.getAttribute("portid"));
        return { ip, ports };
    });

    // Insert table into Word
    await Word.run(async (context) => {
        const table = context.document.body.insertTable(
            hosts.length + 1,
            2,
            Word.InsertLocation.start,
            [["IP Address", "Open Ports"]]
        );

        hosts.forEach((h, i) => {
            table.getCell(i + 1, 0).value = h.ip;
            table.getCell(i + 1, 1).value = h.ports.join(", ");
        });

        await context.sync();
    });
};
