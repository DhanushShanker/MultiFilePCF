import { IInputs, IOutputs } from "./generated/ManifestTypes";

import * as XLSX from "xlsx"; // Excel file handling
import * as mammoth from "mammoth"; // Word file handling
import { renderAsync } from "docx-preview";

interface UploadedFile {
    id: string;        // Unique identifier for the file (e.g., annotation ID)
    filename: string;  // Name of the file
    content: string; // Base64-encoded file content
    mimeType: string; // MIME type of the file (e.g., "application/pdf", "image/jpeg")
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
    private fileIdMap: Map<string, string> = new Map<string, string>();

    constructor() {}

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this.context = context;
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

    
    // Ensure the parent container has position relative
    this.container.style.position = 'relative'; // Apply position: relative to the parent container

    // Create the file preview container
    const filePreviewContainer = document.createElement('div');
    filePreviewContainer.id = "file-preview-container";
    filePreviewContainer.style.display = 'none';
    filePreviewContainer.style.position = 'fixed'; // The preview container will be fixed to the viewport
    filePreviewContainer.style.top = '50%'; // Center vertically (50% of the viewport height)
    filePreviewContainer.style.left = '50%'; // Center horizontally (50% of the viewport width)
    filePreviewContainer.style.transform = 'translate(-50%, -50%)'; // Adjust for exact center
    filePreviewContainer.style.width = '80%'; // Adjust width for a rectangular shape
    filePreviewContainer.style.height = '80%'; // Adjust height for a rectangular shape
    filePreviewContainer.style.backgroundColor = 'rgba(255, 255, 255, 0.95)'; // Semi-transparent background
    filePreviewContainer.style.zIndex = '10000'; // Make sure it overlaps other content
    filePreviewContainer.style.padding = '20px'; // Padding for content inside the preview container
    filePreviewContainer.style.overflowY = 'auto'; // Allow scrolling if content exceeds height
    filePreviewContainer.style.boxShadow = '0 6px 12px rgba(0, 0, 0, 0.3)'; // Optional: Add shadow for better visibility
    filePreviewContainer.style.borderRadius = '10px'; // Slight rounding for rectangle edges
    filePreviewContainer.innerHTML = `
        <h3 style="margin: 0;">File Preview</h3>
        <button id="close-preview-button" 
            style="position: absolute; 
                   top: 15px; 
                   right: 20px; 
                   background-color: #f44336; 
                   color: white; 
                   border: none; 
                   border-radius: 5px; 
                   padding: 5px 10px; 
                   cursor: pointer;">Close Preview</button>
    <div id="file-preview-content" style="margin-top: 20px;"></div>
`;

    // Append the preview container to the body (not this.container)
    document.body.appendChild(filePreviewContainer);

    // Handle the close button functionality
    this.closePreviewButton = document.getElementById("close-preview-button") as HTMLButtonElement;
    if (this.closePreviewButton) {
        this.closePreviewButton.addEventListener("click", () => this.closePreview());
}

    const accountId = this.context.parameters.accountId.raw;
    if (accountId) {
        this.retrieveFiles(accountId);  // Retrieve files associated with the account
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
    
                const reader = new FileReader();
    
                // Read the file content as ArrayBuffer
                reader.onload = async () => {
                    try {
                        const fileContent = reader.result as ArrayBuffer;
                        const base64FileContent = this.base64ArrayBuffer(fileContent);
    
                        console.log(`Uploading file: ${file.name}`);
                        await this.uploadFile(file.name, file.type, base64FileContent, accountId);
                    } catch (error) {
                        console.error(`Error uploading file (${file.name}):`, error);
                        alert(`Failed to upload file: ${file.name}`);
                    }
                };
    
                // Handle errors while reading the file
                reader.onerror = (e) => {
                    console.error(`Error reading file (${file.name}):`, e);
                    alert(`Failed to read file: ${file.name}`);
                };
    
                reader.readAsArrayBuffer(file); // Trigger file read
            });
    
            // Notify that output has changed if needed
            this.notifyOutputChanged();
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
        const deleteButton = this.createFileOptionButton("Delete", () => this.deleteFile(file,fileItem));

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
    
    private async previewWord(file: File): Promise<void> {
        const container = document.getElementById("file-preview-content");
        if (!container) {
            console.error("Error: Element with ID 'file-preview-content' not found.");
            return; // Exit early if the container is null
        }
    
        try {
            const arrayBuffer = await file.arrayBuffer();
            await renderAsync(arrayBuffer, container);
        } catch (error: unknown) {
            console.error("Error rendering document:", error);
            container.innerHTML = "<p>Error rendering document</p>";
        }
    }

    private closePreview(): void {
        const previewContainer = document.getElementById('file-preview-container');
        if (previewContainer) {
            previewContainer.style.display = 'none'; // Hide the preview container
        }
    }

    private async deleteFile(file: File, fileItem: HTMLDivElement): Promise<void> {
        try {
            // Retrieve the file ID from the map
            const fileId = this.fileIdMap.get(file.name);
    
            if (!fileId) {
                console.error("File ID not found. Cannot delete from server.");
                return;
            }
    
            // Delete the file from the server (Notes entity)
            await this.context.webAPI.deleteRecord("annotation", fileId);
            console.log(`File deleted from server: ${fileId}`);
        // Remove the file from the uploadedFiles array
        const index = this.uploadedFiles.indexOf(file);
        if (index > -1) {
            this.uploadedFiles.splice(index, 1);
        }
    
        // Remove the file item from the UI
        this.fileList.removeChild(fileItem);
    
        console.log(`File deleted: ${file.name}`);
    
        // Optionally, notify output change if necessary
        this.notifyOutputChanged();
    }catch (error) {
        console.error("Error deleting file from server:", error);
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

    private async uploadFile(
        filename: string,
        filetype: string,
        base64FileContent: string,
        accountId: string
    ): Promise<void> {
        const annotationEntity = {
            documentbody: base64FileContent,
            filename: filename,
            mimetype: filetype,
            //"objectid_account@odata.bind": `/accounts(${accountId})`,
            "objectid_ats_job_seeker@odata.bind": `/ats_job_seekers(${accountId})`,
            subject: "Uploaded File",
        };
 
        try {
            const response = await this.context.webAPI.createRecord("annotation", annotationEntity);
            console.log("File uploaded successfully as annotation:", response.id);
        } catch (error) {
            console.error("Error creating annotation record:", error);
            throw error;
        }
    }
    
    private base64ArrayBuffer(arrayBuffer: ArrayBuffer): string {
        const bytes = new Uint8Array(arrayBuffer);
        let binary = '';
        for (let i = 0; i < bytes.byteLength; i++) {
            binary += String.fromCharCode(bytes[i]);
        }
        return btoa(binary);
    }

    private decodeBase64(base64: string): ArrayBuffer {
        const binaryString = atob(base64);  // Decode base64 to binary string
        const len = binaryString.length;
        const arrayBuffer = new ArrayBuffer(len);
        const uint8Array = new Uint8Array(arrayBuffer);
    
        for (let i = 0; i < len; i++) {
            uint8Array[i] = binaryString.charCodeAt(i);
        }
    
        return arrayBuffer;
    }

    private async retrieveFiles(accountId: string): Promise<void> {
        if (!accountId) {
            console.error("Account ID is null or undefined. Cannot retrieve files.");
            return;
        }
    
        try {
            console.log("Retrieving files for account:", accountId);
    
            // Replace with your actual API call logic to fetch files
            const retrievedFiles: UploadedFile[] = await this.fetchFilesFromServer(accountId);
    
            // Simulate `File` objects for each retrieved file and display them
            retrievedFiles.forEach((file) => {
                const fileContent = file.content ? this.decodeBase64(file.content) : null;
                if (fileContent) {
                const simulatedFile = new File([fileContent], file.filename || "Unknown file",{ type: file.mimeType });
                this.addFileToList(simulatedFile);
                this.fileIdMap.set(simulatedFile.name, file.id);
            } else {
                console.error("File content is missing or invalid.");
            }
            });
    
        } catch (error) {
            console.error("Error retrieving files:", error);
        }
    }
    
    private async fetchFilesFromServer(accountId: string): Promise<UploadedFile[]> {
        const query = `?$filter=_objectid_value eq ${accountId}`;
        try {
            const result = await this.context.webAPI.retrieveMultipleRecords("annotation", query);
    
            if (result.entities.length > 0) {
                console.log("Retrieved files:", result.entities);
    
                // Map entities to the UploadedFile type
                return result.entities.map((entity) => ({
                    id: entity["annotationid"],  // Assuming annotationid is the unique file identifier
                    filename: entity["filename"] || "Unknown file",
                    content: entity["documentbody"] || "", // Assuming `documentbody` contains the file content
                    mimeType: entity["mimetype"] || "application/octet-stream", // Default to generic binary stream if MIME type is unavailable
                    // Add additional fields here if necessary
                }));
            } else {
                console.log("No files found for the given account.");
                return [];
            }
        } catch (error) {
            console.error("Error fetching files from server:", error);
            throw error; // Re-throw the error to handle it in the calling function
        }
    }
}
