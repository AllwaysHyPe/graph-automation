# GraphUserPhotoSync-Automation

## Overview
This script automates **user photo updates** in **Microsoft Entra ID (Azure AD)** using **Microsoft Graph API**.  
It retrieves **all users** first, **stores them in an ordered hash table**, finds **matching photos in a local directory** (such as an **Azure Arc-enabled server**), and updates their profile photos—without requiring the `Microsoft.Graph` PowerShell module.

### Key Benefits
- **Minimizes Graph API calls** to reduce consumption and improve execution speed.
- **Runs in an Azure Automation Account with a Hybrid Worker via Azure Arc**, avoiding cloud sandbox restrictions.
- **Authentication via a Managed Identity** aka no passwords to manage.
- **Processes users efficiently** by handling all retrievals before performing updates.

---

## How It Works

### **1. Retrieves All Users from Microsoft Graph**
- Connects using a **Managed Identity**.
- Queries **all enabled users** (excluding guests).
- Stores user details in an **ordered hash table** for **fast lookups and reduced API calls**.

### **2. Matches Users to Local Photos**
- Reads the **/Photos/InProgress/** directory from a **local server** (such as an Azure Arc-enabled machine).
- Creates a **lookup table of filenames** (`displayName.jpg`).
- Identifies **users who have a matching photo**.

### **3. Uploads Photos to Microsoft Graph**
- Uses **`Invoke-RestMethod`** to efficiently **PATCH profile photos**.
- Sends **raw binary image data** (no base64 encoding required).
- **Minimizes API calls** by only processing users **who actually need an update**.
- Logs **successes and failures** for easy debugging.

### **4. Moves Successfully Processed Photos**
- After a **successful upload**, moves the photo to **/Photos/Completed/**.
- Keeps the `/InProgress/` folder **clean** and **organized**.

---

## Setup & Prerequisites

### **Required Components**
1. **Azure Automation Account**
   - Must be running on a **Hybrid Worker deployed via Azure Arc**.
   - Must use PowerShell 7 in Automation Account.

2. **Managed Identity with Graph API Permissions**
   - Requires:
     - `User.ReadWrite.All` (to update profile photos)
     - `User.Read.All` (to fetch user details)

3. **Local Photo Directory**
   - **Source Folder:** `/Photos/InProgress/`
   - **Completed Folder:** `/Photos/Completed/`
   - Filenames **must match user display names** (e.g., `Jane Doe.jpg`).
