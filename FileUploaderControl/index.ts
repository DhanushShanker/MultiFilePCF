import { IInputs, IOutputs } from "./generated/ManifestTypes";

import * as XLSX from "xlsx"; // Excel file handling
import * as mammoth from "mammoth"; // Word file handling

interface FileRecord {
    new_filesid: string; // GUID of the file
    new_filename: string; // File name
    new_filecontent?: string; // Base64-encoded content (optional if not retrieved)
    new_mimetype?: string; // MIME type (optional if not retrieved)
    new_accountid?: string; // Associated account ID (optional if not retrieved)
    createdon: string; // Date created
}

export class FileUploaderControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private container: HTMLDivElement;
    private fileInput: HTMLInputElement;
    private fileList: HTMLDivElement;
    private chooseFilesButton: HTMLButtonElement;
    private closePreviewButton: HTMLButtonElement | null = null; 
    private notifyOutputChanged: () => void;
    

    private uploadedFiles: File[] = []; // To hold the list of uploaded files.

    private context: ComponentFramework.Context<IInputs>;

    constructor() {}

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this.container = container;
        this.notifyOutputChanged = notifyOutputChanged;

        const fileUploaderContainer = document.createElement("div");
        fileUploaderContainer.id = "file-uploader-container";

        // Create and append elements.
        this.fileInput = document.createElement("input");
        this.fileInput.type = "file";
        this.fileInput.id = "file-input";
        this.fileInput.multiple = true;
        this.fileInput.style.display = "none"; // Hide the file input

        // Listen for file selection change
        this.fileInput.addEventListener("change", this.handleFileUpload.bind(this));
        
        const fileLabel = document.createElement("label");
        fileLabel.textContent = "Upload Files";
        fileLabel.classList.add("upload-files-label");
        this.container.appendChild(fileLabel);

        // Create the "Choose Files" button
        this.chooseFilesButton = document.createElement("button");
        this.chooseFilesButton.textContent = "Choose Files";
        this.chooseFilesButton.id = "choose-files-button";
        this.chooseFilesButton.addEventListener("click", this.triggerFileInput.bind(this));

        // Create the file list display container
        this.fileList = document.createElement("div");
        this.fileList.id = "file-list";

        

        // Append elements to the container
        this.container.appendChild(this.chooseFilesButton);
        this.container.appendChild(this.fileInput);
        this.container.appendChild(this.fileList);

        // Create the file preview container
        const filePreviewContainer = document.createElement('div');
        filePreviewContainer.id = "file-preview-container";
        filePreviewContainer.style.display = 'none';
        filePreviewContainer.innerHTML = `
            <h3>File Preview</h3>
            <div id="file-preview-content"></div>
            <button id="close-preview-button">Close Preview</button>
        `;
        this.container.appendChild(filePreviewContainer);

        this.closePreviewButton = document.getElementById("close-preview-button") as HTMLButtonElement;
    if (this.closePreviewButton) {
        this.closePreviewButton.addEventListener("click", () => this.closePreview());
    }
    const accountId = context.parameters.accountId.raw;
    if (accountId) {
        this.retrieveFilesForAccount(accountId).then((files) => {
            files.forEach((file) => {
                const fileObj = new File([], file.new_filename, { type: file.new_mimetype || "" });
                this.addFileToList(fileObj);
            });
        });
    } else {
        console.warn("No account ID provided. Skipping file retrieval.");
    }

    }

    

    private triggerFileInput(): void {
        console.log("Button clicked, triggering file input.");
        this.fileInput.click();
    }

    private handleFileUpload(event: Event): void {
        const input = event.target as HTMLInputElement;
        if (input.files) {
            Array.from(input.files).forEach((file) => {
                this.uploadedFiles.push(file);
                this.addFileToList(file);

                const accountId = this.context.parameters.accountId.raw;
                if (!accountId) {
                    console.error("Account ID is null or undefined. File upload aborted.");
                    return;
                }
                this.saveFile(file, accountId).then((fileId) => {
                    if (fileId) {
                        console.log("File saved successfully, File ID: ", fileId);
                    } else {
                        console.log("Failed to save the file");
                    } 
            });

            // Notify that output has changed if needed.
            this.notifyOutputChanged();
        });
    }
}

    private addFileToList(file: File): void {
        const fileItem = document.createElement("div");
        fileItem.className = "file-item";
        fileItem.textContent = file.name;

        // Create options for each file
        const fileOptions = document.createElement("div");
        fileOptions.className = "file-options";
        
        const downloadButton = this.createFileOptionButton("Download", () => this.downloadFile(file));
        const previewButton = this.createFileOptionButton("Preview", () => this.previewFile(file));
        const deleteButton = this.createFileOptionButton("Delete", () => this.deleteFile(file));

        fileOptions.appendChild(downloadButton);
        fileOptions.appendChild(previewButton);
        fileOptions.appendChild(deleteButton);

        fileItem.appendChild(fileOptions);
        this.fileList.appendChild(fileItem);
    }

    private createFileOptionButton(text: string, onClick: () => void): HTMLButtonElement {
        const button = document.createElement("button");
        button.className = "file-option-button";
        button.textContent = text;
        button.addEventListener("click", onClick);
        return button;
    }

    private downloadFile(file: File): void {
        const url = URL.createObjectURL(file);
        const link = document.createElement("a");
        link.href = url;
        link.download = file.name;
        link.click();
        URL.revokeObjectURL(url); // Clean up the object URL
    }

    private previewFile(file: File): void {
        const filePreviewContent = document.getElementById("file-preview-content");
        const filePreviewContainer = document.getElementById("file-preview-container");

        if (filePreviewContainer && filePreviewContent) {
            filePreviewContainer.style.display = 'block'; // Show the preview container

            const fileType = file.type;

            // Check file type and handle preview accordingly
            if (fileType === "application/pdf") {
                // For PDF files
                filePreviewContent.innerHTML = `<embed src="${URL.createObjectURL(file)}" width="100%" height="500px" />`;
            } else if (fileType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || fileType === "application/vnd.ms-excel") {
                // For Excel files, use SheetJS (xlsx library)
                this.previewExcel(file);
            } else if (fileType === "application/vnd.openxmlformats-officedocument.wordprocessingml.document") {
                // For Word files
                this.previewWord(file);
            } else if (fileType.startsWith("image/")) {
                // For image files
                const imageUrl = URL.createObjectURL(file);
                filePreviewContent.innerHTML = `<img src="${imageUrl}" alt="Image Preview" style="max-width: 100%; max-height: 500px;" />`;
            } else {
                // For other types (text files, etc.)
                const reader = new FileReader();
                reader.onload = (e) => {
                    const content = e.target?.result as string;
                    filePreviewContent.innerHTML = `<pre>${content}</pre>`;
                };
                reader.readAsText(file);
            }
        }
    }

    private previewExcel(file: File): void {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = e.target?.result;
            if (data) {
                // Parse Excel file
                const workbook = XLSX.read(data, { type: 'array' });
                const previewContainer = document.getElementById("file-preview-content");
    
                if (previewContainer) {
                    // Clear previous content
                    previewContainer.innerHTML = '';
    
                    // Create tabs for sheet navigation
                    const sheetTabs = document.createElement('div');
                    sheetTabs.id = 'sheet-tabs';
                    sheetTabs.style.marginBottom = '10px';
    
                    // Add each sheet as a tab
                    workbook.SheetNames.forEach((sheetName, index) => {
                        const tabButton = document.createElement('button');
                        tabButton.textContent = sheetName;
                        tabButton.style.marginRight = '5px';
                        tabButton.style.padding = '5px 10px';
                        tabButton.style.cursor = 'pointer';
                        tabButton.style.border = '1px solid #ccc';
                        tabButton.style.backgroundColor = index === 0 ? '#e0e0e0' : '#f9f9f9';
    
                        // On click, render the corresponding sheet
                        tabButton.onclick = () => {
                            document.querySelectorAll('#sheet-tabs button').forEach((btn) => {
                                (btn as HTMLElement).style.backgroundColor = '#f9f9f9';
                            });
                            tabButton.style.backgroundColor = '#e0e0e0';
                            this.renderSheet(workbook, sheetName);
                        };
    
                        sheetTabs.appendChild(tabButton);
                    });
    
                    // Append tabs and render the first sheet by default
                    previewContainer.appendChild(sheetTabs);
                    this.renderSheet(workbook, workbook.SheetNames[0]);
                }
            }
        };
        reader.onerror = (error) => {
            console.error("Error reading file:", error);
        };
        reader.readAsArrayBuffer(file);
    }
    
    private renderSheet(workbook: XLSX.WorkBook, sheetName: string): void {
        const worksheet = workbook.Sheets[sheetName];
        const previewContainer = document.getElementById("file-preview-content");
    
        if (worksheet && previewContainer) {
            // Generate HTML table with row and column headers
            const html = this.generateTableWithHeaders(worksheet);
            const sheetContent = document.getElementById('sheet-content');
            if (sheetContent) {
                sheetContent.remove();
            }
    
            const tableContainer = document.createElement('div');
            tableContainer.id = 'sheet-content';
            tableContainer.innerHTML = html;
            previewContainer.appendChild(tableContainer);
    
            // Apply custom styling
            this.applyExcelStyling();
        }
    }
    
    private generateTableWithHeaders(worksheet: XLSX.WorkSheet): string {
        // Convert sheet to JSON array
        const jsonData: (string | number | boolean | null)[][]  = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
        // Generate table headers (A, B, C...)
        const columnHeaders = `<tr><th></th>${jsonData[0]
            .map((_, colIndex) => `<th>${String.fromCharCode(65 + colIndex)}</th>`)
            .join('')}</tr>`;
    
        // Generate rows with row numbers (1, 2, 3...)
        const rows = jsonData
            .map(
                (row, rowIndex) =>
                    `<tr><th>${rowIndex + 1}</th>${row
                        .map((cell) => `<td>${cell || ''}</td>`)
                        .join('')}</tr>`
            )
            .join('');
    
        // Combine headers and rows into a table
        return `<table id="excel-preview-table">${columnHeaders}${rows}</table>`;
    }
    
    private applyExcelStyling(): void {
        const table = document.getElementById('excel-preview-table');
        if (table) {
            // General table styles
            table.style.borderCollapse = 'collapse';
            table.style.width = '100%';
            table.style.tableLayout = 'auto';
    
            // Table header styles
            const headers = table.querySelectorAll('th');
            headers.forEach((header) => {
                header.style.backgroundColor = '#f0f0f0'; // Excel-like color
                header.style.border = '1px solid #d0d0d0';
                header.style.padding = '5px';
                header.style.textAlign = 'center';
                header.style.fontWeight = 'bold';
            });
    
            // Table cell styles
            const cells = table.querySelectorAll('td');
            cells.forEach((cell) => {
                cell.style.border = '1px solid #d0d0d0';
                cell.style.padding = '5px';
            });
    
            // Font family to match Excel's default
            table.style.fontFamily = 'Calibri, Arial, sans-serif';
        }
    }
    

    private previewWord(file: File): void {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = e.target?.result as ArrayBuffer;
            const arrayBuffer = new Uint8Array(data);
            // Assuming you're using Mammoth.js (or similar) to parse Word document
            mammoth.convertToHtml({ arrayBuffer: data }).then((result) => {
                document.getElementById("file-preview-content")!.innerHTML = result.value;
                return result;
            })
            .catch((error) => {
                console.error("Error converting Word document:", error);
                document.getElementById("file-preview-content")!.innerHTML = "<p>Error rendering document</p>";
                throw error;
            });
        };
        reader.onerror = (error) => {
            console.error("Error reading file:", error);
            document.getElementById("file-preview-content")!.innerHTML = "<p>Error reading file</p>";
        };
        reader.readAsArrayBuffer(file);
    }

    private closePreview(): void {
        const previewContainer = document.getElementById('file-preview-container');
        if (previewContainer) {
            previewContainer.style.display = 'none'; // Hide the preview container
        }
    }

    // Implement the deleteFile method
    private deleteFile(file: File): void {
        const fileIndex = this.uploadedFiles.indexOf(file);
        if (fileIndex !== -1) {
            // Remove the file from the array
            this.uploadedFiles.splice(fileIndex, 1);
            // Re-render the file list
            this.fileList.innerHTML = ''; // Clear the file list
            this.uploadedFiles.forEach((file) => {
                this.addFileToList(file); // Re-add files to the list
            });
            // Notify that output has changed if needed.
            this.notifyOutputChanged();
        }
    }

    // Implement the updateView method
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        // Update the control based on the new data from the inputs
        // This could involve re-rendering the file list or handling other input changes.
        console.log("updateView called");
        // You can add additional logic here to refresh or re-render the control
    }

    public getOutputs(): IOutputs {
        return {
            FileData: this.uploadedFiles.map((file) => file.name).join(", ")
        };
    }

    public destroy(): void {
        // Cleanup the control.
        this.fileInput.removeEventListener("change", this.handleFileUpload.bind(this));
        this.chooseFilesButton.removeEventListener("click", this.triggerFileInput.bind(this));
    }

    private async saveFile(file: File, accountId: string): Promise<string | null> {
        if (!accountId || accountId.length === 0) {
            console.error("No valid accountId provided");
            return null;
        }
    
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
    
            reader.onload = async (event) => {
                try {
                    const fileContent = event.target?.result as ArrayBuffer;
                    const base64Content = this.base64ArrayBuffer(fileContent);
    
                    // Prepare the file data for the Files entity
                    const fileData = {
                        "new_accountid@odata.bind": `/accounts(${accountId})`, // Associate the file with the account
                        "new_filename": file.name, // File name
                        "new_filecontent": base64Content, // Base64 encoded file content
                        "new_mimetype": file.type, // File MIME type
                    };
    
                    // Use Xrm.Utility.getGlobalContext().getClientUrl() to get the client URL
                    const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl();
    
                    const response = await fetch(`${clientUrl}/api/data/v9.2/new_files`, {
                        method: "POST",
                        headers: {
                            "Content-Type": "application/json",
                            "OData-MaxVersion": "4.0",
                            "OData-Version": "4.0",
                            "Accept": "application/json",
                        },
                        body: JSON.stringify(fileData),
                    });
    
                    if (response.ok) {
                        const result = await response.json();
                        console.log("File saved successfully with ID:", result.new_filesid);
                        resolve(result.new_filesid); // Resolve with the File record ID
                    } else {
                        console.error("Failed to save file:", await response.text());
                        resolve(null); // Resolve with null on failure
                    }
                } catch (error) {
                    console.error("Error saving file:", error);
                    resolve(null); // Resolve with null on error
                }
            };
    
            reader.onerror = () => {
                console.error("Error reading the file");
                resolve(null); // Resolve with null on file read error
            };
    
            reader.readAsArrayBuffer(file);
        });
    }
    
    private base64ArrayBuffer(arrayBuffer: ArrayBuffer): string {
        const bytes = new Uint8Array(arrayBuffer);
        let binary = '';
        for (let i = 0; i < bytes.byteLength; i++) {
            binary += String.fromCharCode(bytes[i]);
        }
        return btoa(binary);
    }
    
    private async retrieveFilesForAccount(accountId: string): Promise<FileRecord[]> {
        if (!accountId || accountId.length === 0) {
            console.error("No valid accountId provided");
            return [];
        }
    
        // Use Xrm.Utility.getGlobalContext().getClientUrl() to get the client URL
        const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl();
    
        try {
            // Query the Files entity for files associated with the account
            const response = await fetch(`${clientUrl}/api/data/v9.2/new_files?$filter=_new_accountid_value eq ${accountId}`, {
                method: "GET",
                headers: {
                    "Accept": "application/json",
                    "OData-MaxVersion": "4.0",
                    "OData-Version": "4.0",
                },
            });
    
            if (response.ok) {
                const result = await response.json();
                console.log("Files retrieved successfully:", result.value);
                return result.value; // Return the list of files
            } else {
                console.error("Failed to retrieve files:", await response.text());
                return [];
            }
        } catch (error) {
            console.error("Error retrieving files:", error);
            return [];
        }
    }
    
}
