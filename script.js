function processFile() {

    const fileInput = document.getElementById("fileInput");
    const file = fileInput.files[0];
    const status = document.getElementById("status");

    if (!file) {
        status.innerHTML = "Please select a file.";
        return;
    }

    // ❗ Show processing message
    document.getElementById("processingMsg").style.display = "block";

    // File validation
    if (!file.name.endsWith(".docx")) {
        document.getElementById("processingMsg").style.display = "none";
        status.innerHTML = "Unsupported file type. Please upload a .docx file.";
        return;
    }

    status.innerHTML = "Processing document...";

    const reader = new FileReader();

    reader.onload = function (event) {
        JSZip.loadAsync(event.target.result).then(function (zip) {

            const docFile = zip.file("word/document.xml");
            const coreFile = zip.file("docProps/core.xml");

            if (!docFile) {
                document.getElementById("processingMsg").style.display = "none";
                status.innerHTML = "Invalid document structure.";
                return;
            }

            docFile.async("string").then(function (xmlText) {

                // Simple text extraction
                let text = xmlText.replace(/<[^>]+>/g, " ");

                // Split text into chunks (simulate RSID analysis)
                let words = text.split(/\s+/);
                let resultHTML = "";

                words.forEach(word => {
                    if (word.length > 10) {
                        resultHTML += `<span style="color:red;">${word} </span>`;
                    } else {
                        resultHTML += `<span style="color:blue;">${word} </span>`;
                    }
                });

                document.getElementById("documentPreview").innerHTML = resultHTML;

                // Metadata
                document.getElementById("fileName").innerText = file.name;
                document.getElementById("fileSize").innerText = (file.size / 1024).toFixed(2) + " KB";
                document.getElementById("fileStatus").innerText = "Processed";

                if (coreFile) {
                    coreFile.async("string").then(function (metaText) {

                        const getTag = (tag) => {
                            const match = metaText.match(new RegExp(`<${tag}[^>]*>(.*?)</${tag}>`));
                            return match ? match[1] : "Not available";
                        };

                        document.getElementById("creator").innerText = getTag("dc:creator");
                        document.getElementById("created").innerText = getTag("dcterms:created");
                        document.getElementById("modified").innerText = getTag("dcterms:modified");

                    });
                }

                // Show result
                document.getElementById("result").classList.remove("hidden");

                // ❗ Hide processing message
                document.getElementById("processingMsg").style.display = "none";

                status.innerHTML = "Analysis complete.";

            });

        }).catch(function () {
            document.getElementById("processingMsg").style.display = "none";
            status.innerHTML = "Error processing file.";
        });
    };

    reader.readAsArrayBuffer(file);
}

function resetApp() {
    document.getElementById("fileInput").value = "";
    document.getElementById("result").classList.add("hidden");
    document.getElementById("status").innerHTML = "";
}
