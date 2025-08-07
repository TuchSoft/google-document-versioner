# Docs Versioning for AppSheet

This Google Apps Script project automates the versioning of Google Docs and synchronizes this data with an AppSheet application. The script adds a unique identifier and version number to the document's footer and tracks all revisions and associated metadata in a connected AppSheet database.

## Table of Contents

  - [Docs Versioning for AppSheet](https://www.google.com/search?q=%23docs-versioning-for-appsheet)
      - [Table of Contents](https://www.google.com/search?q=%23table-of-contents)
      - [Background](https://www.google.com/search?q=%23background)
      - [How It Works](https://www.google.com/search?q=%23how-it-works)
      - [Features](https://www.google.com/search?q=%23features)
      - [How to Use](https://www.google.com/search?q=%23how-to-use)
      - [Step 1: AppSheet Setup](https://www.google.com/search?q=%23step-1-appsheet-setup)
      - [Step 2: Google Apps Script Setup](https://www.google.com/search?q=%23step-2-google-apps-script-setup)
      - [Step 3: Initial Publication to Your Organization](https://www.google.com/search?q=%23step-3-initial-publication-to-your-organization)
      - [How to Update the Published Add-on](https://www.google.com/search?q=%23how-to-update-the-published-add-on)
      - [Authors](https://www.google.com/search?q=%23authors)
      - [License](https://www.google.com/search?q=%23license)

## Background

This script was developed to streamline document management and version control. By linking Google Docs directly to an AppSheet database, it provides a simple way to create, track, and manage document versions without manual data entry.

## How It Works

The script is triggered from a Google Docs custom menu. When you select an action, it automatically generates or updates a unique document ID and a version number in the document's footer. It then captures key details about the document and its revision, including a hash of the content, and pushes this data to your AppSheet application. A minor version bumps the version number by `0.1` (e.g., 1.0 to 1.1), while a major version increments it by `1.0` (e.g., 1.1 to 2.0). A key feature is that **each saved version is "published" privately within your organization**, making a PDF of that specific revision permanently available and accessible via a unique URL.

## Features

  - **Automated Versioning**: Add minor or major version numbers with a single click.
  - **Unique Document IDs**: Assigns a unique, timestamp-based ID to each new document.
  - **AppSheet Integration**: Automatically sends document and version data to a connected AppSheet application.
  - **Revision Tracking**: Captures and logs each version's unique hash, revision ID, and a PDF export link.
  - **Permanent PDF Archive**: **Publishes each saved revision, creating a stable and shareable PDF link for every version.** This is done privately for your organization.
  - **Document Details**: Extracts document names, descriptions (from headings), and notes for each version.
  - **Folder-based Context**: Automatically links new documents to customers and projects based on the folder structure in Google Drive.
  - **Dashboard Links**: Provides menu items to open the document's details or the main dashboard in AppSheet.

## How to Use

1.  **Open or Create a Google Doc**: Start with any Google Doc.
2.  **Access the Menu**: The custom menu **`Versions`** will appear in the toolbar.
3.  **Run the Script**:
      - Select **`Minor version (+0.1)`** to create a minor version.
      - Select **`Major version (+1.0)`** to create a major version.
4.  **Confirm and Add Notes**: A popup will ask for confirmation and allow you to add optional notes for the new version.
5.  **View in AppSheet**: Use the **`Open Doc Details`** or **`Open Dashboard`** menu items to navigate to your AppSheet application.

-----
## How to install

### Step 1: AppSheet Setup

First, you need to create the AppSheet database that the script will connect to.

1.  **Create a New AppSheet App**: Go to [AppSheet](https://www.appsheet.com) and create a new app. You can start with a blank app or from a Google Sheet.

2.  **Define the Tables**: You need to create four tables with the following minimum columns. These tables are the core of your database.

    #### 1\. Customers

    | Column Name | Type | Key |
    | :--- | :--- | :--- |
    | **Row ID** | `Text` | Primary Key |
    | **Name** | `Text` | |

    #### 2\. Projects

    | Column Name | Type | Key |
    | :--- | :--- | :--- |
    | **Row ID** | `Text` | Primary Key |
    | **Name** | `Text` | |

    #### 3\. Documents

    | Column Name | Type | Key |
    | :--- | :--- | :--- |
    | **Row ID** | `Text` | Primary Key |
    | **Id** | `Text` | Alternate Key |
    | **Name** | `Text` | |
    | **Description** | `LongText` | |
    | **Date** | `DateTime` | |
    | **Customer** | `Ref` | |
    | **Project** | `Ref` | |

    #### 4\. Documents Versions

    | Column Name | Type | Key |
    | :--- | :--- | :--- |
    | **Row ID** | `Text` | Primary Key |
    | **Name** | `Text` | |
    | **Date** | `DateTime` | |
    | **Notes** | `LongText` | |
    | **Document** | `Ref` | |
    | **Hash** | `Text` | |
    | **Revision** | `Text` | |
    | **DocId** | `Text` | |
    | **Url** | `URL` | |
    | **Pdf** | `URL` | |

3.  **Get AppSheet Credentials**:

      - Go to your AppSheet app's **Manage \> Integrations \> API** page.
      - Find your **Application Access Key** and **App ID**. You will need these to configure the Apps Script.

-----

### Step 2: Google Apps Script Setup

Next, you will create the Apps Script project and link it to your AppSheet app.

1.  **Create a New Apps Script Project**: Go to [Google Apps Script](https://script.google.com) and create a **`New project`**.
2.  **Copy and Paste the Code**: Copy the entire script provided and paste it into the `Code.gs` file of your new project, replacing any existing code.
3.  **Configure the Constants**: At the top of the script, update the following constants with the information from your AppSheet app:
      - `APPSHEET_APP_ID`: Your AppSheet app ID.
      - `APPSHEET_ACCESS_KEY`: Your AppSheet access key.
4.  **Enable Advanced Services**: The script uses the Google Drive API to manage document revisions.
      - In the Apps Script editor, go to the left-hand menu and click **`+`** next to **`Services`**.
      - Select **`Drive API`** and click **`Add`**.
5.  **Save the Project**: Click the save icon to save your script.

-----

### Step 3: Initial Publication to Your Organization

To make this add-on available to all users in your Google Workspace organization, you need to publish it through the Google Workspace Marketplace.

1.  **Manifest Configuration**: Before publishing, you must enable the Google Workspace Marketplace SDK in your Google Cloud project.
      - In the Apps Script editor, go to **`Project Settings`** (the gear icon).
      - Under **`Google Cloud Platform (GCP) Project`**, note the Project Number and click on the link to open the GCP console.
      - In the GCP console, go to **`APIs & Services` \> `Library`**.
      - Search for and enable the **`Google Workspace Marketplace SDK`**.
2.  **Deploy as a Web App**:
      - In the Apps Script editor, go to `Deploy` \> `New deployment`.
      - Select **`Add-on`** as the type.
      - Fill out the **`Deployment settings`**:
          - **Version**: `New version`.
          - **Description**: A brief description of this initial deployment.
          - Click **`Deploy`**.
3.  **Configure for the Marketplace**:
      - After deployment, you will be directed to the Google Cloud project page for the add-on.
      - Go to the **`App Configuration`** tab.
      - Fill in the necessary details, including:
          - **App Integration**: Select **`Google Docs Add-on`**.
          - **Docs Add-on Script Version**: Enter the version number you just deployed.
      - Click **`Save Draft`**.
4.  **Publish to Organization**:
      - Go to the **`Store Listing`** tab.
      - Provide a title, description, and graphics for your add-on.
      - Under **`Visibility`**, choose **`My Organization`** to make it available internally.
      - Click **`Publish`** at the bottom of the page. Your add-on will now be available for installation by users in your organization.

-----

## How to Update the Published Add-on

Once the add-on is published, you can update it with new features or bug fixes.

1.  **Update the Script**: Ensure the `APPVERSION` constant at the top of the code is updated.
2.  **Deploy a New Version**:
      - In the Apps Script editor, go to `Deploy` \> `Manage Deployments`.
      - Select your existing deployment, click the pencil icon, and choose `New version`.
      - Click **`Deploy`**.
3.  **Google Cloud Console**:
      - Go to the Google Cloud Console for your project.
      - Navigate to the **Google Workspace Marketplace SDK** page.
      - In the **`App Configuration`** tab, update the **"Docs Add-on Script Version"** to match your new deployment version.
      - Click **`Save Draft`**.
4.  **Publish**:
      - Go to the **`Store Listing`** tab.
      - Click **`Publish`** at the bottom of the page. The updated version will now be available to your organization's users.

-----

## Authors

  - Tuchsoft [https://tuchsoft.com](https://tuchsoft.com)
  - Mattiabonzi [mattia@tuchsoft.com](mailto:mattia@tuchsoft.com)

## License

This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).
