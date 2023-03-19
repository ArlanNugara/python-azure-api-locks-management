# Getting Started
To work on this project, you should have the git client and an editor (VS Code is free and capable)

If this is your first time using git, you will need to tell it your name and email address. This can be done using the following two commands (making the obvious changes)

`git config --global user.name "Your Name"`

`git config --global user.email you@example.com`

# Introduction

Coming Soon..

# Requirements in Azure

* A Storage Account
* A containers in Storage Accounts
* Storage Account Access Keys

# Setup Azure DevOps

## Import GitHub repository

The Azure tag Management GitHub repository is public hence no authorization is required. Clone URL: `https://github.com/ArlanNugara/python-azure-api-locks-management.git`

Check here for [Importing a repo](https://docs.microsoft.com/en-us/azure/devops/repos/git/import-git-repository?view=azure-devops#import-into-a-new-repo)

## Create Azure DevOps Service Connection

For the pipeline to authenticate and connect to Azure we need to create an Azure DevOps Service Connection which basically is a Service Principal (Application)

### Create Service Principle

* In Azure Portal navigate to **Azure Active Directory**
* Click on **App registrations**
* Click on **New registration**
* Name the application (e.g. **Tags_MGMT_SP**)
* Click **Register**
* Once the App registration is created, in the **Overview**' note the **Application ID**
* Under **Manage** click on **Certificates & Secrets**
* Click on **New client secret**
* Provide a description and choose the expiry time and click **Add** button
* A new client secret is generated, copy the secret´s value

### Assign Permission to Service Principle

* In Azure portal navigate to **Subscriptions**
* Select the subscription at which Azure Tag Management will run
* Go to **Access Control (IAM)**
* Click on **Add Role Assignment**
* Select **User Access Administrator** and **Reader** and click on **Next** button
* Select the Service Principle name created in above step and click on **Next** button
* Click on **Review + Assign** button

Note: Provide permission at Management Group if you have more than one subscription for ease of configuration

### Create Azure DevOps Service Connection

* Click on **Project settings** from left panel bottom
* Click on **Service Connections** under **Pipelines**
* Click on **New service connection** and select the service or connection type as **Azure Resource Manager** and click on **Next** button
* Select **Service principal (manual)** and click on **Next** button
* Select the Environment as **Azure Cloud** 
* Select the Scope Level (You may select Subscription or Management Group based on your requirements)
* Provide **Subscription ID** / **Management Group ID** and **Subscription Name** / **Management Group Name** based on your selection in previous step
* Under **Authentication** in the field **Service Principal Id** enter the **Application ID** that was created in previous step
* For the **Credential** select **Service principal key**, in the field **Service principal key** enter the secret key that was generated in previous step
* For **Tenant ID** enter the Tenant Id.
* Click on **Verify** button to verify connection. In case of error check on above details and ensure there is no typo mistakes
* Under **Details** provide the Service Connection name. Note the Service Connection Name.
* Click on **Verify and save** button

## Create Azure DevOps Variable Group

* Click on **Library** under **Piplines** in Azure DevOps Project
* Click on **+ Variable Group** button on top
* Provide a name for the variable group. Note the variable group name
* Click on **+Add** button and add the following variables with values (Details below)
    * ARM_CLIENT_ID
    * ARM_CLIENT_SECRET
    * ARM_TENANT_ID
    * ARM_SUBSCRIPTION_ID
    * ARM_ACCESS_KEY
    * SA_NAME
    * CONTAINER_NAME
* Click on **Save** button

|Variable Name|Description|
|:-----------:|:---------:|
|ARM_CLIENT_ID|The application ID of the Service Principle Created|
|ARM_CLIENT_SECRET|The secret value of the Service Principle Created|
|ARM_TENANT_ID|The Tenant ID|
|ARM_SUBSCRIPTION_ID|The Subscription ID|
|ARM_ACCESS_KEY|Storage Account Access Keys|
|SA_NAME|Storage Account Name|
|CONTAINER_NAME|Container Name where Reports will be uploaded|

Note: Repeat the above steps if you have more than one subscription.

## Create the Pipeline

### Edit Pipeline variables File

* Open **Tags.variables.yml** file in **.pipelines** directory
* Replace `<VariableGroup>` with the variable group name created
* Replace `<ServiceConnection>` with the service connection name created
* Replace `<ReportsFileNamePrefix>` with the prefix of file name you want for reports without spaces(e.g `All-Locks` or `Locks-Report` etc)
* Save the file

### Create the Pipelines

* Select **Pipelines** from Azure DevOps project
* Click on **New Pipeline**
* Select **Azure Repos Git**
* Select the repository you imported
* Select **Existing Azure Pipelines YAML file**
* Select the repository branch and select the Tags Report Pipeline file path - **/.pipelines/resource-lock-report.yml**
* Click on **Continue** button
* Click on **Save** button from **Run** Dropdown menu.

Note: Repeat the step for Locks Update Pipeline. Choose the file path as **/.pipelines/resource-lock-update.yml**.

Note : Add 2 variables in the update pipeline named **FILENAME** and **DATE**

# Get Azure Locks

This process creates an excel file which contains locks of all resources, resource groups for a given subscription. The excel file is then pushed to a storage container in Azure using the same Azure DevOps Pipeline which generates the excel file.

## Get Locks process summary

This is the general idea how the process works within the Pipeline and Python

* Start the pipeline selecting the subscription name which triggers the python script with few arguments.
* Login to Azure using the Service Principle credentials to generate the API tokens.
* Query Azure API to get information about a scope i.e Subscription or Resource group or Resource.
* Format and inserts the data into Excel Sheet.
* Upload the excel file to Azure Storage Account Container.

![image](./images/get-tags.png)

## Running the Pipeline

* Click on **Pipelines**
* Select the Azure Tags Report Pipeline
* Click **Run pipeline**

Note: For first time user the Pipeline will require authorization for accessing the Variable Group. Authorize the same while runnign the pipeline for the first time.

## Get the Report

* Go to the storage account and click on **Containers** from left side panel
* Click on the container created for locks
* Download the Excel File

# Update Locks

This process updates locks for any scope i.e Subscription or Resource group or Resource. It uses the excel file created in [get azure locks](#get-azure-locks) to run. You can modify the excel file values as per your requirement.

## Update Locks process summary

Coming Soon..