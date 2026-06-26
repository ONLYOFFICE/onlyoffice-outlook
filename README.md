# ONLYOFFICE add-in for Outlook

This add-in allows users to view, edit, and create office files directly in [Microsoft Outlook](https://www.microsoft.com/en-us/microsoft-365/outlook/email-and-calendar-software-microsoft-outlook) using ONLYOFFICE Docs.

## Features ⭐️

* View and edit documents, spreadsheets, PDFs, and presentations attached to email messages.
* Create new files (DOCX, XLSX, PPTX) directly from the email composition window.
* Download and delete attachments without leaving Outlook.

### Supported formats

**For viewing:**

- **WORD**: DOC, DOCM, DOCX, DOT, DOTM, DOTX, EPUB, FB2, FODT, GDOC, HML, HTM, HTML, HWP, HWPX, MD, MHT, MHTML, ODT, OTT, PAGES, RTF, STW, SXW, TXT, WPS, WPT, XML
- **CELL**: CSV, ET, ETT, FODS, GSHEET, NUMBERS, ODS, OTS, SXC, TSV, XLS, XLSB, XLSM, XLSX, XLT, XLTM, XLTX
- **SLIDE**: DPS, DPT, FODP, GSLIDES, KEY, ODG, ODP, OTP, POT, POTM, POTX, PPS, PPSM, PPSX, PPT, PPTM, PPTX, SXI
- **PDF**: DJVU, DOCXF, OFORM, OXPS, PDF, XPS
- **DIAGRAM**: VSDM, VSDX, VSSM, VSSX, VSTM, VSTX

**For editing:**

- **WORD**: DOCM, DOCX, DOTM, DOTX
- **CELL**: XLSB, XLSM, XLSX, XLTM, XLTX
- **SLIDE**: POTM, POTX, PPSM, PPSX, PPTM, PPTX
- **PDF**: PDF

## Requirements

- ONLYOFFICE Docs (Document Server)

## Installing ONLYOFFICE Docs

To be able to work with office files within Outlook, you will need an instance of [ONLYOFFICE Docs](https://www.onlyoffice.com/office-suite.aspx). You can install the self-hosted version of the editors or opt for ONLYOFFICE Docs Cloud which doesn't require downloading and installation.

### Self-hosted editors

You can install free Community version of ONLYOFFICE Docs or scalable Enterprise Edition.

To install free Community version, use [Docker](https://github.com/onlyoffice/Docker-DocumentServer) (recommended) or follow [these instructions](https://helpcenter.onlyoffice.com/docs/installation/docs-community-install-ubuntu.aspx) for Debian, Ubuntu, or derivatives.

To install Enterprise Edition, follow the instructions [here](https://helpcenter.onlyoffice.com/docs/installation/enterprise).

Community Edition vs Enterprise Edition comparison can be found [here](#onlyoffice-docs-editions).

### ONLYOFFICE Docs Cloud

To get ONLYOFFICE Docs Cloud, get started [here](https://www.onlyoffice.com/docs-registration.aspx).

## Add-in installation 📥

### Install from Microsoft AppSource

1. Go to [appsource.microsoft.com](https://appsource.microsoft.com/).
2. Search for **ONLYOFFICE** and select the ONLYOFFICE Docs add-in for Outlook.
3. Click **Get it now** and follow the on-screen instructions to add it to your Outlook account.

### Install via manifest file (sideloading)

1. Open Outlook and click the **Get Add-ins** button (or go to **More options** → **Get Add-ins**).
2. In the add-ins dialog, select **My add-ins** in the left panel.
3. Scroll down to the **Custom add-ins** section and click **Add a custom add-in** → **Add from file**.
4. Select the `manifest.xml` file and click **Open**.
5. Click **Install** to confirm.

## Add-in configuration ⚙️

After installation, open the add-in settings by clicking the settings icon in the ONLYOFFICE panel. On the settings page, configure:

* **ONLYOFFICE Docs address**: The URL of your installed ONLYOFFICE Docs instance (e.g. `https://your-docs-server.com`).

* **ONLYOFFICE Docs secret key**: Enables JWT to protect your documents from unauthorized access (further information can be found [here](https://helpcenter.onlyoffice.com/docs/installation/docs-configure-jwt.aspx)).

## Add-in usage

Open any email message in Outlook — the ONLYOFFICE panel appears in the reading pane and lists all supported attachments. When composing a new message, the panel is accessible from the toolbar.

### Viewing and editing files 📝

Click the **eye** icon to open a file in view mode, or the **pencil** icon to open it in edit mode. The editor launches in a pop-up window powered by ONLYOFFICE Docs. After editing, saving the document updates the email attachment automatically.

### Creating files 📄

In compose mode, click **Document**, **Spreadsheet**, or **Presentation** to create a new file. Enter a file name in the form that appears and click **Create**. The new file opens in the editor immediately and is added as an email attachment.

### Downloading and deleting files

Click the **download** icon next to any attachment to save it locally. In compose mode, click the **delete** icon to remove an attachment from the message.

## Development

1. Clone project from the GitHub repository:
```
git clone https://github.com/ONLYOFFICE/onlyoffice-outlook
```

2. Install the project dependencies:
```
npm install
```

3. Start the development server:
```
npm run dev-server
```

4. Build the project for production:
```
npm run build
```

5. Validate the add-in manifest:
```
npm run validate
```

6. Start debugging (sideloads the add-in into the configured Office host):
```
npm start
```

7. Stop debugging:
```
npm stop
```

## ONLYOFFICE Docs editions

ONLYOFFICE offers different versions of its online document editors that can be deployed on your own servers.

**ONLYOFFICE Docs** packaged as Document Server:

* Community Edition 🆓 (`onlyoffice-documentserver` package) – Perfect for small teams and personal use.
* Enterprise Edition 🏢 (`onlyoffice-documentserver-ee` package) – Designed for businesses with advanced features & support.

The table below will help you make the right choice.

| Pricing and licensing | Community Edition | Enterprise Edition |
| ------------- | ------------- | ------------- |
| | [Get it now](https://www.onlyoffice.com/download-community.aspx?utm_source=github&utm_medium=cpc&utm_campaign=GitHubOutlook#docs-community)  | [Start Free Trial](https://www.onlyoffice.com/download.aspx?utm_source=github&utm_medium=cpc&utm_campaign=GitHubOutlook#docs-enterprise)  |
| Cost  | FREE  | [Go to the pricing page](https://www.onlyoffice.com/docs-enterprise-prices.aspx?utm_source=github&utm_medium=cpc&utm_campaign=GitHubOutlook)  |
| Simultaneous connections | up to 20 maximum  | As in chosen pricing plan |
| Number of users | up to 20 recommended | As in chosen pricing plan |
| License | GNU AGPL v.3 | Proprietary |
| **Support** | **Community Edition** | **Enterprise Edition** |
| Documentation | [Help Center](https://helpcenter.onlyoffice.com/docs/installation/community) | [Help Center](https://helpcenter.onlyoffice.com/docs/installation/enterprise) |
| Standard support | [GitHub](https://github.com/ONLYOFFICE/DocumentServer/issues) or paid | One or three years support included |
| Premium support | [Contact us](mailto:sales@onlyoffice.com) | [Contact us](mailto:sales@onlyoffice.com) |
| **Services** | **Community Edition** | **Enterprise Edition** |
| Conversion Service                | + | + |
| Document Builder Service          | + | + |
| **Interface** | **Community Edition** | **Enterprise Edition** |
| Tabbed interface                       | + | + |
| Dark theme                             | + | + |
| 125%, 150%, 175%, 200% scaling         | + | + |
| White Label                            | - | - |
| Integrated test example (node.js)      | + | + |
| Mobile web editors                     | - | +* |
| **Plugins & Macros** | **Community Edition** | **Enterprise Edition** |
| Plugins                           | + | + |
| Macros                            | + | + |
| **Collaborative capabilities** | **Community Edition** | **Enterprise Edition** |
| Two co-editing modes              | + | + |
| Comments                          | + | + |
| Built-in chat                     | + | + |
| Review and tracking changes       | + | + |
| Display modes of tracking changes | + | + |
| Version history                   | + | + |
| **Document Editor features** | **Community Edition** | **Enterprise Edition** |
| Font and paragraph formatting   | + | + |
| Object insertion                | + | + |
| Adding Content control          | + | + |
| Editing Content control         | + | + |
| Layout tools                    | + | + |
| Table of contents               | + | + |
| Navigation panel                | + | + |
| Mail Merge                      | + | + |
| Comparing Documents             | + | + |
| **Spreadsheet Editor features** | **Community Edition** | **Enterprise Edition** |
| Font and paragraph formatting   | + | + |
| Object insertion                | + | + |
| Functions, formulas, equations  | + | + |
| Table templates                 | + | + |
| Pivot tables                    | + | + |
| Data validation           | + | + |
| Conditional formatting          | + | + |
| Sparklines                   | + | + |
| Sheet Views                     | + | + |
| **Presentation Editor features** | **Community Edition** | **Enterprise Edition** |
| Font and paragraph formatting   | + | + |
| Object insertion                | + | + |
| Transitions                     | + | + |
| Animations                      | + | + |
| Presenter mode                  | + | + |
| Notes                           | + | + |
| Slide Master	                  | + | + |
| **Form creator features** | **Community Edition** | **Enterprise Edition** |
| Adding form fields           | + | + |
| Form preview                    | + | + |
| Saving as PDF                   | + | + |
| **PDF Editor features**      | **Community Edition** | **Enterprise Edition** |
| Text editing and co-editing	  | + | + |
| Work with pages (adding, deleting, rotating) | + | + |
| Inserting objects (shapes, images, hyperlinks, etc.) | + | + |
| Text annotations (highlight, underline, cross out) | + | + |
| Comments                        | + | + |
| Freehand drawings               | + | + |
| Form filling                    | + | + |
| | [Get it now](https://www.onlyoffice.com/download-community.aspx?utm_source=github&utm_medium=cpc&utm_campaign=GitHubOutlook#docs-community)  | [Start Free Trial](https://www.onlyoffice.com/download.aspx?utm_source=github&utm_medium=cpc&utm_campaign=GitHubOutlook#docs-enterprise)  |

\* If supported by DMS.

## Need help? Feedback & Support 💡

In case of technical problems, the best way to get help is to submit your issues [here](https://github.com/ONLYOFFICE/onlyoffice-outlook/issues). Alternatively, you can contact ONLYOFFICE team via [community.onlyoffice.com](https://community.onlyoffice.com) or [feedback.onlyoffice.com](https://feedback.onlyoffice.com/forums/966080-your-voice-matters).
