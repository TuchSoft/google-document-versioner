# Docs Versioning for Google Documents

This Google Apps Script project automates the versioning of Google Docs and synchronizes this data with an AppSheet application. The script adds a unique identifier and version number to the document's footer and tracks all revisions and associated metadata in a connected AppSheet database.

## Table of Contents

  - [Docs Versioning for AppSheet](https://www.google.com/search?q=%23docs-versioning-for-appsheet)
      - [Table of Contents](https://www.google.com/search?q=%23table-of-contents)
      - [Background](https://www.google.com/search?q=%23background)
      - [How It Works](https://www.google.com/search?q=%23how-it-works)
      - [Features](https://www.google.com/search?q=%23features)
      - [How to Use](https://www.google.com/search?q=%23how-to-use)
      - [How to Publish to an Organization](https://www.google.com/search?q=%23how-to-publish-to-an-organization)
      - [AppSheet Setup](https://www.google.com/search?q=%23appsheet-setup)
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

## How to Publish to an Organization

To deploy this script as a Google Workspace Add-on for your organization, follow these steps:

1.  **Update the Script**: Ensure the `APPVERSION` constant at the top of the code is updated.
2.  **Deploy**:
      - In the Apps Script editor, go to `Deploy` \> `Manage Deployments`.
      - Select your existing deployment, click the pencil icon, and choose `New version`.
      - Click `Deploy`.
3.  **Google Cloud Console**:
      - Go to the Google Cloud Console for your project.
      - Navigate to the **APIs & Services** dashboard and find **Google Workspace Marketplace SDK**.
      - In the **App Configuration** tab, update the **"Docs Add-on Script Version"**.
      - Click **`Save Draft`**.
4.  **Publish**:
      - Go to the **Store Listing** tab.
      - Click **`Publish`** at the bottom of the page.

## AppSheet Setup

The script requires a connected AppSheet application with four tables. Here is the minimum data structure required for each table to function correctly:

### 1\. Customers

| Column Name | Type | Key |
| :--- | :--- | :--- |
| **Row ID** | `Text` | Primary Key |
| **Name** | `Text` | |

### 2\. Projects

| Column Name | Type | Key |
| :--- | :--- | :--- |
| **Row ID** | `Text` | Primary Key |
| **Name** | `Text` | |

### 3\. Documents

| Column Name | Type | Key |
| :--- | :--- | :--- |
| **Row ID** | `Text` | Primary Key |
| **Id** | `Text` | Alternate Key |
| **Name** | `Text` | |
| **Description** | `LongText` | |
| **Date** | `DateTime` | |
| **Customer** | `Ref` | |
| **Project** | `Ref` | |

### 4\. Documents Versions

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

## Authors

  - Tuchsoft [https://tuchsoft.com](https://tuchsoft.com)
  - Mattiabonzi [mattia@tuchsoft.com](mailto:mattia@tuchsoft.com)

## License

This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).
