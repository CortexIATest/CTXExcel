# CTXExcel
Cortex Subtasks which interact with Microsoft Excel.


## Table of Contents
1) [Dependencies](#dependencies)
    * [Cortex Version](#cortex-version)
    * [OCIs](#ocis)
    * [Files](#files)
    * [Other](#other)
1) [Installation](#installation)
1) [How to use](#how-to-use)
1) [How you can contribute](#how-you-can-contribute)
1) [FAQs](#faqs)
1) [The boring bits](#the-boring-bits)
    * [Versioning](#versioning)
    * [Licensing](#licensing)

## Dependencies
### Cortex Version
This version of the CTXExcel module was developed in Cortex v6.3.0. Some functionality may not be available in earlier verions of Cortex.

### OCIs
The CTXExcel module requires the following Cortex OCIs:
* PowerShell

### Files
The CTXExcel module requires the following files:
* [CTXExcel.studiopkg](https://github.com/CortexIATest/CTXExcel/blob/master/CTXExcel-V2.2.studiopkg)

### Other
The CTXExcel module has the following additional requirements which will be explained later:
* [PowerShell v5](#powershell-v5) to be installed on the application server
* [PSExcel](#psexcel) PowerShell module installed
* [ImportExcel](#importexcel) PowerShell module installed

## Installation
Details of the installation can be found in the [CTXExcel Deployment Plan](https://github.com/CortexIATest/CTXExcel/blob/master/CTXExcel%20Deployment%20Plan%20-%20v2.2.docx), an overview of the install is given below.

### PowerShell v5
For the CTXExcel module to work, a requirement is that Powershell version 5 is installed on the Cortex Server. This can be checked by opening Powershell and running the following command:

`$PSVersionTable.PSVersion`

A similar output should be displayed as below:

```
Major  Minor  Build  Revision
-----  -----  -----  --------
5      1      16299  64
```

If the major version is 5 or greater, move on to [PSExcel](#psexcel). 

If the major version is less than 5, perform the following steps:
* On the Cortex server where Powershell version 5 will be installed, navigate to this [link](
https://www.microsoft.com/en-us/download/details.aspx?id=50395&tduid=(162666df8fd7d1ab0239724a9bec1eca)(266696)(1503186)(61836X1384699Xf82af593098584c381b4505006d7472d)())
* Click the ‘Download’ button
* Select and download the version required for the server where Powershell is being installed
* Run the installer and follow the instructions provided [here](https://docs.microsoft.com/en-us/powershell/wmf/5.1/install-configure)

### PSExcel
To install the PSExcel PowerShell module, perform the following steps:
* Run Powershell as an administrator 
* Execute the following query:

`Install-Module PSExcel`

* A dialogue box may come up and ask if you trust this repositry, click 'Yes'

### ImportExcel
To install the ImportExcel PowerShell module, perform the following steps:
* Run Powershell as an administrator 
* Execute the following query:

`Install-Module ImportExcel`

* A dialogue box may come up and ask if you trust this repositry, click 'Yes'

### Import CTXExcel Module
To import the CTXExcel Module to your Cortex server, perform the following steps:
* Connect to Cortex Server
* Download [CTXExcel.studiopkg](https://github.com/CortexIATest/CTXExcel/blob/master/CTXExcel-V2.2.studiopkg) on the Cortex application server
* Navigate to Settings &rarr; Studio Import
* Select 'Browse' and select the [CTXExcel.studiopkg](https://github.com/CortexIATest/CTXExcel/blob/master/CTXExcel-V2.2.studiopkg) downloaded
* Import the module
* Ensure the studio authorisations are set up for the relevant user groups

## How to use
A detailed Low-Level Design (LLD) document has been provided with instructions on how to use the flows/subtasks, available [here](https://github.com/CortexIATest/CTXExcel/blob/master/CTXExcel%20-%20LLD%20-%20v2.2.docx). Configuration of subtask inputs and outputs are detailed in notes on the subtask workspace.

## How you can contribute
While the CTXExcel subtasks contain a plethora of functionality already, we are always looking to add more. If there is any functionality you desire, please get in touch.

Unfortunately, we cannot offer pull requests at this time. 

## FAQs
For any questions which are not answered here, please contact the author responsible for the functionality you are using.

* How is this module so good? 
   * It was made by the Waterboy himself, that's how. 

## The boring bits
### Versioning
The CTXExcel module has the following versions, starting with the most recent:
* CTXExcel v2.2 - Author: Aleo Yakas, email: testemail@testing.com
  *  Add Conditional Formatting

---
* CTXExcel v2.1 - Author: Aleo Yakas, email: testemail@testing.com
  *  Create Pivot Charts

---
* CTXExcel v2.0 - Author: Aleo Yakas, email: testemail@testing.com
  *  Create Pivot Tables

---
* CTXExcel v1.0 - Author: Aleo Yakas, email: testemail@testing.com
  *  Create Workbook
  *  Create Worksheet
  *  Insert Headers
  *  Insert Rows
  *  Set Cell
  *  Set Cell Where

### Licensing
All functionality within this module is licensed under the [MIT license](https://opensource.org/licenses/mit-license.php). 
