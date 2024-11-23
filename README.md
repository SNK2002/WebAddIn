Excel Web Add-in
This project is an Excel Web Add-in that allows users to view and interact with PDF files directly within Excel. The add-in is built using HTML, CSS, and JavaScript, and it integrates with Excel to enhance productivity.
Features
PDF Viewer: Select and view PDF files within the Excel task pane.
Excel Integration: Seamlessly integrates with Excel to provide additional functionality.
Modern UI: Utilizes Office UI Fabric for a consistent and modern user interface.
Installation
Prerequisites
Excel 2016 or later
Node.js and npm (for development)
Git (for version control)
Setup
Clone the Repository:
2. Install Dependencies:
If your project has any npm dependencies, install them:
3. Sideload the Add-in:
Open Excel.
Go to "Insert" > "My Add-ins" > "Manage My Add-ins" > "Upload My Add-in".
Select the AddIn.xml manifest file.
Usage
1. Open Excel.
2. Load the Add-in: Navigate to the "Insert" tab and select your add-in.
3. Select a PDF: Use the "Select PDF" button to choose a PDF file from your local system.
4. View PDF: The selected PDF will be displayed in the task pane.
Development
File Structure
AddInWeb/: Contains the web assets for the add-in.
Home.html: Main HTML file for the add-in.
Home.css: Styles for the add-in.
Home.js: JavaScript logic for the add-in.
AddIn/AddInManifest/: Contains the manifest file for the add-in.
Scripts
Build: Compile and prepare the add-in for deployment.
Start: Run a local development server.
Configuration
.gitignore: Specifies files and directories to be ignored by Git.
package.json: Contains metadata and dependencies for the project.
Contributing
Fork the repository.
2. Create a new branch: git checkout -b feature/your-feature-name.
Commit your changes: git commit -m 'Add some feature'.
Push to the branch: git push origin feature/your-feature-name.
Open a pull request.
License
This project is licensed under the MIT License.
Contact
For questions or support, please contact rushabhcmu@gmail.com.
---
Feel free to customize this README further to fit your project's specific details and requirements!
