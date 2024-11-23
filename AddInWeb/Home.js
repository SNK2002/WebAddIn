let messageBanner;
let selectionBox;
let isSelecting = false;
let isFrozen = false;

// Set the worker source for PDF.js
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.10.377/pdf.worker.min.js';

// Initialization when Office JS and JQuery are ready.
Office.onReady(() => {
    $(() => {
        const element = document.querySelector('.MessageBanner');
        messageBanner = new components.MessageBanner(element);
        messageBanner.hideBanner();

        $('#pdf-file-input').on('change', handleFileSelect);
        document.getElementById('snippet-button').addEventListener('click', toggleSnippetMode);
        document.getElementById('confirm-snippet-button').addEventListener('click', confirmSnippet);
    });
});

function toggleSnippetMode() {
    if (isFrozen) {
        resetSelection();
    } else {
        enableSnippetMode();
    }
}

function enableSnippetMode() {
    const pdfCanvas = document.getElementById('pdf-canvas');
    selectionBox = document.createElement('div');
    selectionBox.style.position = 'absolute';
    selectionBox.style.border = '2px dashed #0078d4';
    selectionBox.style.zIndex = '1000';
    document.body.appendChild(selectionBox);

    let startX, startY;

    const onMouseDown = (e) => {
        if (isFrozen) return;
        startX = e.pageX;
        startY = e.pageY;
        isSelecting = true;
        selectionBox.style.left = `${startX}px`;
        selectionBox.style.top = `${startY}px`;
    };

    const onMouseMove = (e) => {
        if (!isSelecting || isFrozen) return;
        const currentX = e.pageX;
        const currentY = e.pageY;
        selectionBox.style.width = `${Math.abs(currentX - startX)}px`;
        selectionBox.style.height = `${Math.abs(currentY - startY)}px`;
        selectionBox.style.left = `${Math.min(currentX, startX)}px`;
        selectionBox.style.top = `${Math.min(currentY, startY)}px`;
    };

    const onMouseUp = () => {
        if (isFrozen) return;
        isSelecting = false;
        isFrozen = true;
        pdfCanvas.removeEventListener('mousedown', onMouseDown);
        pdfCanvas.removeEventListener('mousemove', onMouseMove);
        pdfCanvas.removeEventListener('mouseup', onMouseUp);
    };

    pdfCanvas.addEventListener('mousedown', onMouseDown);
    pdfCanvas.addEventListener('mousemove', onMouseMove);
    pdfCanvas.addEventListener('mouseup', onMouseUp);
}

function confirmSnippet() {
    if (!selectionBox) return;

    const pdfCanvas = document.getElementById('pdf-canvas');
    const rect = selectionBox.getBoundingClientRect();
    html2canvas(pdfCanvas, {
        x: rect.left,
        y: rect.top,
        width: rect.width,
        height: rect.height
    }).then(canvas => {
        const link = document.createElement('a');
        link.href = canvas.toDataURL();
        link.download = 'snippet.png';
        link.click();

        resetSelection();
    }).catch(error => {
        console.error('Error capturing snippet:', error);
    });
}

function resetSelection() {
    if (selectionBox) {
        document.body.removeChild(selectionBox);
        selectionBox = null;
    }
    isFrozen = false;
}

function handleFileSelect(event) {
    const file = event.target.files[0];
    if (file && file.type === 'application/pdf') {
        const fileReader = new FileReader();
        fileReader.onload = function() {
            const typedarray = new Uint8Array(this.result);

            pdfjsLib.getDocument(typedarray).promise.then(pdf => {
                return pdf.getPage(1);
            }).then(page => {
                const scale = 1.5;
                const viewport = page.getViewport({ scale: scale });
                const canvas = document.getElementById('pdf-canvas');
                const context = canvas.getContext('2d');
                canvas.height = viewport.height;
                canvas.width = viewport.width;

                const renderContext = {
                    canvasContext: context,
                    viewport: viewport
                };
                page.render(renderContext);
                $('#pdf-viewer-container').show();
            }).catch(error => {
                console.error('Error loading PDF: ', error);
            });
        };
        fileReader.readAsArrayBuffer(file);
    } else {
        console.error('Please select a valid PDF file.');
    }
}
