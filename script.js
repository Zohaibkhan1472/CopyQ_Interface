let selectedFile = null;

const dropArea = document.getElementById("dropArea");

dropArea.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropArea.classList.add("dragover");
});

dropArea.addEventListener("dragleave", () => {
    dropArea.classList.remove("dragover");
});

dropArea.addEventListener("drop", (e) => {
    e.preventDefault();
    dropArea.classList.remove("dragover");

    selectedFile = e.dataTransfer.files[0];
    document.getElementById("fileInput").files = e.dataTransfer.files;
});

function processFile() {

    const fileInput = document.getElementById("fileInput");
    const status = document.getElementById("status");
    const result = document.getElementById("result");

    const file = selectedFile || fileInput.files[0];

    if (!file) {
        status.innerText = "Please upload a document to proceed.";
        status.style.color = "red";
        return;
    }

    if (file.name.endsWith(".pdf")) {
        status.innerText = "PDF format is not supported. Please upload a .docx document.";
        status.style.color = "red";
        return;
    }

    if (!file.name.endsWith(".docx")) {
        status.innerText = "Invalid file type. Only .docx documents are supported.";
        status.style.color = "red";
        return;
    }

    status.innerText = "Processing document. Please wait...";
    status.style.color = "black";

    const reader = new FileReader();

    reader.onload = function(event) {

        JSZip.loadAsync(event.target.result).then(function(zip) {

            const core = zip.file("docProps/core.xml").async("string");
            const doc = zip.file("word/document.xml").async("string");

            return Promise.all([core, doc]);

        }).then(function([coreXML, docXML]) {

            status.innerText = "Analysis completed successfully.";
            status.style.color = "green";

            result.classList.remove("hidden");

            const parser = new DOMParser();
            const coreDoc = parser.parseFromString(coreXML, "text/xml");

            function get(tag) {
                const el = coreDoc.getElementsByTagName(tag)[0];
                return el ? el.textContent : "N/A";
            }

            document.getElementById("fileName").innerText = file.name;
            document.getElementById("fileSize").innerText = file.size + " bytes";
            document.getElementById("fileStatus").innerText = "Valid document";

            document.getElementById("creator").innerText = get("dc:creator");
            document.getElementById("created").innerText = get("dcterms:created");
            document.getElementById("modified").innerText = get("dcterms:modified");

            const docParsed = parser.parseFromString(docXML, "text/xml");
            const paragraphs = docParsed.getElementsByTagName("w:p");

            let output = "";

            for (let p of paragraphs) {
                let text = "";
                const runs = p.getElementsByTagName("w:t");

                for (let r of runs) {
                    text += r.textContent + " ";
                }

                if (text.length > 200) {
                    output += `<p style="color:red;">${text}</p>`;
                } else {
                    output += `<p style="color:blue;">${text}</p>`;
                }
            }

            document.getElementById("documentPreview").innerHTML = output;

        }).catch(() => {
            status.innerText = "Error processing document.";
            status.style.color = "red";
        });
    };

    reader.readAsArrayBuffer(file);
}

function resetApp() {
    location.reload();
}