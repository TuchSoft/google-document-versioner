/**
 * To publish the script, follow these steps:
 * !! UPDATE THE "APPVERSION" CONSTANT !!
 * Run Deployment > Manage Deployments
 * Select "Untitled"
 * Click on the pencil icon
 * Select "New version" under version
 * Click "Deploy"
 * Go to https://console.cloud.google.com/apis/api/appsmarket-component.googleapis.com/googleapps_sdk?project={your-project}
 * "App Configuration" tab
 * Increase the version under "Docs Add-on Script Version"
 * Click "Save Draft"
 * Go to the "Store Listing" tab
 * Click "Publish" at the bottom
 */

/** !! UPDATE THIS CONSTANT !! */
const APPVERSION = 41;

const TEMPLATE_FOLDER_ID = "x-xxxxxxxxxxxxxxxxxxxx-xxxxxxxxxx";
const APPSHEET_APP_ID = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx';
const APPSHEET_ACCESS_KEY = 'xx-xxxx-xxxx-xxxx-xxxx-xxxx-xxxx-xxxx-xxxx';

// AppSheet table names
const TABLE_PROJECTS = "Projects";
const TABLE_CUSTOMERS = 'Customers';
const TABLE_DOCUMENTS = "Documents";
const TABLE_VERSIONS = "Documents%20Versions";
const APPSHEET_API_BASE_URL = `https://api.appsheet.com/api/v2/apps/${APPSHEET_APP_ID}/tables/`;
const APPSHEET_DASHBOARD_URL = `https://www.appsheet.com/start/${APPSHEET_APP_ID}`;

// Fixed search strings and file names
const FOLDER_PROJECTS = "Progetti";
const FOOTER_RIF_SEARCH_STRING = "TuchSoft OÜ - rif";
const EMPTY_DOCUMENT_NAMES = ["Untitled document"];


function onOpen() {
    const ui = DocumentApp.getUi();
    Logger.log("App version: " + APPVERSION);

    ui.createMenu('Versions')
        .addItem('Minor version (+0.1)', 'minor')
        .addItem('Major version (+1.0)', 'major')
        .addSeparator()
        .addItem('Open Doc Details', 'openDocDetails')
        .addItem('Open Dashboard', 'openAppSheetDashboard')
        .addToUi();
}


function minor() {
    addId();
}

function major() {
    addId(true);
}

function addId(major = false, docId = null) {
    Logger.log("App version: " + APPVERSION);
    const now = new Date();
    const itDate = now.toLocaleDateString('it-IT');
    const doc = docId ? DocumentApp.openById(docId) : DocumentApp.getActiveDocument();
    const footer = doc.getFooter();
    const hash = getDocumentHash(doc);

    Logger.log("Starting addId function.");
    Logger.log("Document title: " + doc.getName());

    if (!footer) {
        alert("No footer found. Exiting.");
        return;
    }

    const existingVersion = get(TABLE_VERSIONS, {
        'Hash': hash
    });
    if (existingVersion && existingVersion.length > 0) {
        if (!confirm("A seemingly identical document already exists, do you want to proceed?'")) {
            Logger.log('Process cancelled');
            return;
        }
    }
    if (EMPTY_DOCUMENT_NAMES.includes(doc.getName()) || listTemplates().includes(doc.getName())) {
        if (!confirm("The document seems to have no name, do you want to proceed?'")) {
            Logger.log('Process cancelled');
            return;
        }
    }

    let document = {
        Id: null,
        Name: doc.getName(),
        Description: getDesc(doc),
        Date: now.toISOString(),
        Customer: null,
    }
    let version = {
        Name: "1.0",
        Date: now.toISOString(),
        Notes: docId ? '' : addNotes(),
        Document: null,
        Hash: hash,
        Revision: null,
        DocId: doc.getId(),
        Url: `https://docs.google.com/document/d/${doc.getId()}/edit`,
        Pdf: ''
    }

    const rif = getRifElement(doc);
    if (rif == null) {
        Logger.log('The document does not contain references, process terminated');
        DocumentApp.getUi().alert("The document does not contain references; it must include the string 'TuchSoft OÜ - rif' in the footer!");
        return;
    }

    const text = rif.getText();
    let search = FOOTER_RIF_SEARCH_STRING;

    if (text.includes(search)) {
        version.Name = "1.0";
        document.Id = genId();
        Logger.log("New document detected. Generating ID: " + document.Id);
    } else {
        let info = parseDoc(text);
        Logger.log("Existing document detected. Parsed info: " + JSON.stringify(info));
        if (info.id === 'undefined') {
            alert('Something went wrong, please check manually.');
            return;
        }
        search = info.search;
        document.Id = info.id;
        version.Name = bumpVersion(info.version, major);
        Logger.log("New version number: " + version.Name);
    }

    // Insert ID into the document
    rif.replaceText(search, `TuchSoft OÜ - ${itDate} - DOC ${document.Id} V${version.Name}`);
    Logger.log("Footer text updated.");

    const props = PropertiesService.getScriptProperties();
    let data = props.getProperty('DATA');
    data = data ? JSON.parse(data) : [];
    data.push([document, version]);
    props.setProperty('DATA', JSON.stringify(data));

    Logger.log("Partial document generated: " + JSON.stringify(document));
    Logger.log("Partial version generated: " + JSON.stringify(version));

    ScriptApp.newTrigger('processNewRevision')
        .timeBased()
        .after(1000)
        .create();
}


function processNewRevision() {
    Logger.log("App version: " + APPVERSION);
    // This function is triggered by 'addId' after a delay.
    var props = PropertiesService.getScriptProperties();
    var data = JSON.parse(props.getProperty('DATA'));

    if (!data || data.length == 0) {
        Logger.log("Nothing to process...");
        return;
    }

    var [document, version] = data.shift();
    props.setProperty('DATA', JSON.stringify(data));

    Logger.log("Processing document: " + JSON.stringify(document));
    Logger.log("Processing version: " + JSON.stringify(version));

    const doc = DocumentApp.openById(version.DocId);
    const path = getPath(doc);

    if (version.Name == '1.0') {
        if (path[0] == FOLDER_PROJECTS) {
            Logger.log("The document is in a customer folder");
            let customer = get(TABLE_CUSTOMERS, {
                Name: path[1]
            });
            if (!customer || customer.length == 0) {
                customer = add(TABLE_CUSTOMERS, {
                    Name: path[1]
                }).Rows[0];
                Logger.log("New customer added: " + JSON.stringify(customer));
            } else {
                customer = customer[0];
            }
            if (!customer) {
                Logger.log("ERROR: something went wrong retrieving or creating the customer");
            } else {
                document.Customer = customer['Row ID'];
                if (path[2] != doc.getName()) {
                    Logger.log("The document is in a project folder");
                    let project = get(TABLE_PROJECTS, {
                        Name: path[2]
                    });
                    if (!project) {
                        //TODO: create project
                    } else {
                        project = project[0];
                    }
                    if (!project) {
                        Logger.log("ERROR: something went wrong retrieving or creating the project");
                    }
                    document.Project = project['Row ID'];
                }
            }
        }
        //Add document
        document = add(TABLE_DOCUMENTS, document).Rows[0];
    } else {
        document = get(TABLE_DOCUMENTS, {
            Id: document.Id
        })[0];
    }

    if (!document) {
        Logger.log("ERROR: missing document");
        return;
    }

    version.Document = document['Row ID'];
    version.Revision = getLastRevision(version.DocId);
    Logger.log('Last revision fetched: ' + version.DocId)
    version.Pdf = publishRevision(version.DocId, version.Revision);

    // Add the new version to AppSheet
    var newVersionAppSheet = add(TABLE_VERSIONS, version).Rows[0];
    Logger.log("Version added to AppSheet: " + JSON.stringify(newVersionAppSheet));

    // Delete the trigger to avoid it running again.
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === 'processNewRevision') {
            ScriptApp.deleteTrigger(triggers[i]);
        }
    }

    Logger.log("Process Finished");
}

// ------------------------------------------------------------------------------------------------------------------



function listTemplates() {

    try {
        const folder = DriveApp.getFolderById(TEMPLATE_FOLDER_ID);
        const files = folder.getFiles();

        let fileNames = [];

        while (files.hasNext()) {
            fileNames.push(files.next().getName());
        }

        if (fileNames.length > 0) {
            Logger.log('Template found: "%s":\n%s', folder.getName(), fileNames.join('\n'));
        } else {
            Logger.log('No files found.');
        }

        return fileNames;

    } catch (e) {
        Logger.log('Error: Could not access template folder with ID "%s". Check ID and permissions. Details: %s', TEMPLATE_FOLDER_ID, e.message);
        return null;
    }
}


function getRifElement(doc) {
    const footer = doc.getFooter();
    if (!footer) {
        alert("No footer found. Exiting.");
        return null;
    }

    for (let i = 0; i < footer.getNumChildren(); i++) {
        let element = footer.getChild(i);

        if (element.getType() == DocumentApp.ElementType.PARAGRAPH) {
            let textElement = element.asParagraph().editAsText();
            let text = textElement.getText();
            Logger.log("Footer paragraph text: " + text);
            if (text.includes("TuchSoft OÜ - ")) {
                return textElement;
            }
        }
    }
    return null;
}

function bumpVersion(ver, major = false) {
    const match = ver.split('.');

    let majorVer = parseInt(match[0]);
    let minorVer = parseInt(match[1]);

    if (major) {
        majorVer++;
        minorVer = 0;
    } else {
        minorVer++;
    }
    return majorVer + "." + minorVer;
}


function getPath(doc) {
    var file = DriveApp.getFileById(doc.getId());
    var path = [];

    function buildPath(file) {
        var parents = file.getParents();
        if (parents.hasNext()) {
            var parent = parents.next();
            buildPath(parent);
            path.push(parent.getName());
        }
    }

    buildPath(file);
    path.push(file.getName());
    path.shift();
    return path;
}

function parseDoc(str) {
    const regex = /TuchSoft\sOÜ\s-\s(\d{2}\/\d{2}\/\d{4})\s-\sDOC\s(.*?)\sV(\d{1,3}\.\d{1,3})/;
    const match = str.match(regex);

    if (match) {
        return {
            date: match[1],
            id: match[2],
            version: match[3],
            search: match[0]
        };
    }

    return null;

}

function genId() {
    // The id is the number of seconds since 01/01/2025 00:00:00
    return Math.floor(Date.now() / 1000) - (60 * 60 * 24 * 365 * 55);
}

function add(db, ...objs) {
    var data = {
        "Action": "Add",
        "Rows": objs
    }
    return post(db, data);
}

function get(db, filter = null) {
    var data = {
        "Action": "Find",
    }
    if (filter) {
        for (let el in filter) {
            data.Properties = {
                Selector: `Filter(${db}, [${el}] = \"${filter[el]}\")`
            };
        }
    }
    return post(db, data);
}


function post(db, data) {
    var url = `${APPSHEET_API_BASE_URL}${db}/Action`;
    var options = {
        'method': 'post',
        'contentType': 'application/json',
        'payload': JSON.stringify(data),
        'headers': {
            'ApplicationAccessKey': APPSHEET_ACCESS_KEY
        }
    };

    Logger.log(`Raw post url: ${url} data: ` + JSON.stringify(data));
    var response = UrlFetchApp.fetch(url, options);
    const responseData = response.getContentText()
    Logger.log("Raw post response: " + responseData);

    return JSON.parse(responseData);
}

function alert(str) {
    Logger.log("ALERT (check the doc page): " + str)
    DocumentApp.getUi().alert(str);
}


function getDesc(doc) {
    var body = doc.getBody();
    var elements = body.getParagraphs();
    let desc = '';

    for (var i = 0; i < elements.length; i++) {
        var element = elements[i];
        if (element.getHeading() == DocumentApp.ParagraphHeading.HEADING1) {
            desc = element.getText();
            break;
        }
    }

    for (var i = 0; i < elements.length; i++) {
        var element = elements[i];
        if (element.getHeading() == DocumentApp.ParagraphHeading.SUBTITLE) {
            desc += ', ' + element.getText();
            break;
        }
    }

    return desc;
}

function getDocumentHash(doc) {
    var hash = sha256(doc.getBody().getText());
    Logger.log("SHA-256 hash of the document: " + hash);
    return hash;
}

function getLastRevision(docId) {
    var allRevisions = [];
    var pageToken = null;
    do {
        var result = Drive.Revisions.list(docId, {
            pageSize: 100,
            pageToken: pageToken,
            fields: 'nextPageToken, revisions(id)'
        });
        if (result.revisions) {
            allRevisions = allRevisions.concat(result.revisions);
        }
        pageToken = result.nextPageToken;
    } while (pageToken);
    console.log(allRevisions);
    if (allRevisions.length > 0) {
        return allRevisions.pop().id;
    }

    return null;
}

function publishRevision(docId, revId) {
    var x = Drive.Revisions.get(docId, revId);
    console.log(x);

    Logger.log("Publishing revision: " + revId + " for doc: " + docId)
    Drive.Revisions.update({
        published: true,
    }, docId, revId);

    var revision = Drive.Revisions.get(docId, revId, {
        fields: 'exportLinks'
    });

    return revision.exportLinks['application/pdf'] || '';
}

function sha256(inputString) {
    return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, inputString)
        .reduce((output, byte) => output + (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, '0'), '');
}

function addNotes() {
    var ui = DocumentApp.getUi();
    Logger.log("MODAL (check the doc page)");
    var prompt = ui.prompt(
        'Add notes for the new version (optional):',
        'Notes',
        ui.ButtonSet.OK_CANCEL
    );

    if (prompt.getSelectedButton() == ui.Button.OK) {
        var note = prompt.getResponseText();
        if (note) {
            Logger.log('User notes: ' + note);
            return note;
        }
    } else {
        Logger.log('User cancelled.');
    }
    return '';
}

function confirm(str) {
    const ui = DocumentApp.getUi();
    Logger.log("MODAL (check the doc page): " + str);
    const response = ui.alert(str, ui.ButtonSet.YES_NO);
    return response === ui.Button.YES;
}

function openUrl(url) {
    var html = HtmlService.createHtmlOutput('<html><script>' +
            'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};' +
            'var a = document.createElement("a"); a.href="' + url + '"; a.target="_blank";' +
            'if(document.createEvent){' +
            '  var event=document.createEvent("MouseEvents");' +
            '  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}' +
            '  event.initEvent("click",true,true); a.dispatchEvent(event);' +
            '}else{ a.click() }' +
            'close();' +
            '</script>'
            +
            '<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="' + url + '" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>' +
            '<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>' +
            '</html>')
        .setWidth(90).setHeight(1);
    DocumentApp.getUi().showModalDialog(html, "Opening ...");
}


function openAppSheetDashboard() {
    Logger.log("App version: " + APPVERSION);
    openUrl(APPSHEET_DASHBOARD_URL);
}

function openDocDetails() {
    Logger.log("App version: " + APPVERSION);
    doc = DocumentApp.getActiveDocument();
    const info = parseDoc(getRifElement(doc).getText());
    if (info.id) {
        const doc = get(TABLE_DOCUMENTS, {
            Id: info.id
        })[0];
        if (doc) {
            openUrl(`${APPSHEET_DASHBOARD_URL}#view=Documents_Detail&row=${doc['Row ID']}`);
        } else {
            alert('Document not registered');
        }
    } else {
        alert('Document without ID');
    }
}
