# DESIGN

This document describes the design that goes behind finding a way to automate inventory management for OCHCA HDP department. We, for now, will use a local method of using Power Automates desktop version. Power Automate has a cloud version that has actions that directly connect to the Atlassian ecosystem, but it seems like there is no cloud access as of now. 

The main issue we face is the lack in being able to create and run code without running into security issues. We need to find a way to remain inside an ecosystem that Zscaler wont create roadblocks in. The main flow thus far is to use Power Automate to run PowerShell scripts to do just about everything. We have used the PowerShell to make the API requests and access the Atlassian ecosystem. Below we will describe the steps we are taking for the building out the Assets Objects side of the project.

Another thing to keep in mind is how this will flow. The big idea behind the flow is, on a defined time basis, whether that is once a week or once a day, the flow triggers and runs. For now the main one that will run is the script reading our source of truth. This source of truth, for now, is an excel file living locally. Any changes should reflect assets which in the future will reflect jira issues.

The next step after figuring out the assets project is to then make it so that we can access/create/modify Jira issues, for now we figure out the assets side of the project.

I have made this repo to share with other interns in an effort to tackle this together. When making a pull request the goal is to be very detailed like I have been in this README.md in order to successfully run the scripts locally to see how the process works, and make meaningful contributions. The path that everyone should have as a baseline is, `Z:\01 Intern Work\Current Interns\YOUR_NAME\project`. You can consider this as a way of having the project locally for your changes and testing.

Inside `notes` there is additional information regarding the structure of this project. Pending implementations for future PR's can be found inside `notes\flowchart.md` at the bottom of the file.

## Create mock data

Start by creating a mock data sample that can be used to simulate having data (inventory) on some file, making changes to the data to hopefully reflect changes on assets. A Power Automate flow called `create_excel` is used for just that. Script can be found in the `scripts` folder in a file called `scripts\create_excel.ps1`.

Line 3 requires changing the path to your local path. You can also create input variables which I will explain below.

## Building the Schema

Having an excel file build a schema from its entries would be a nightmare so we think the best path would be to instead manually build the schema in a structure that would make the most sense with what our actual data presents and then work on some sort of mapping rule to go from excel to Assets. There is room for future improvement here. In a perfect world we can have a schema built from an excel file? Or is it better to keep the excel file as simple as possible and only ever add complexity in Assets? Input is welcome.

Current schema can be found at, `https://ochca.atlassian.net/jira/assets/object-schema/5?mode=attribute&typeId=69&view=list`, ensure you are logged into your work account.

## Save object types and attributes in a JSON

After building the schema the next step is to extract all the object type and attributes into a JSON to then use to build an excel file that will help begin the mapping process. This file can be found at `scripts\get_json_from_assets.ps1`.

This is where we first begin needing an api token, you create one here: `https://id.atlassian.com/manage-profile/security/api-tokens`.

The script is divided into three scripts divided by "%%%%%%%%%%%%%%%%" lines. I created these scripts as three different actions inside a Power Automate flow. Creating a flow in Power Automate is straight forward. After creating a flow you can then run actions. The `Run PowerShell script` action can be found under the Scripting tab on the left.

The first script sends a request to obtain your cloud id. When right clicking an action you can change the `Variables produced` name to `cloud_id` that way later actions can use those variables.

The second script/action requires you to create additional input variables. On the right of the Power Automate flow screen, there is an `Input` tab where you can create those variables. Hit the $\oplus$ button to create variables. You will need to create an `api_token` and `email` variables for this action/script. This script ultimately outputs your workspace id. Again you will need to modify the name of `Variables produced`.

The final script is what outputs a json of the structure of the schema in assets. A `folder_path` input variable will be need to be created. Ensure this variable has the path to your local project. The outputted json will be found at this file path and will be used by future scripts.

## JSON to a excel that begins mapping process

In a new flow you will need to create two additional variables, `folder_path` and `json_file_path`. These paths should again point to `Z:\01 Intern Work\Current Interns\YOUR_NAME\project`, the json path should additionally have the name of the file `\assets_schema_5_export.JSON` which is the output of the previous script. This flow consists of only one action that ultimately creates the excel file that will help with the mapping process. The structure of this created excel file is of two sheets. The `ObjectTypes` sheet is a sheet each row represents an object type found in the schema. Object names, their ID's, and other MetaData related to the objects can be found on this sheet. The second sheet, `Attributes`, is a sheet where each row represents an attribute found in the schema. Each row has additional MetaData that will help with future files.

The final excel holding all the data is called `assets_schema_5_flattened.xlsx`. The is and additional `assets_flatten_debug.log` file that is created to help with debugging in case there are issues in creating the excel file.

This script is all contained in one file and was ran on Power Automate inside one flow using one action. This can be found under `script\json_to_flattened.ps1`.

# SUMMARY OF LAST SESSION 06/17

Today I finished debugging the `json_to_flattened` file and have successfully created the excel that will help construct the actual mapping file.

Next step is to finish task 3 which is the implementation of the mapping file that will help tell future scripts where the each source of truth points to in assets. Last left of in the main file in chatgippity.

Additionally this repo was created to help catch others up to speed with the current progress of the project. The goal is to read this md file and have an intuitive idea of the structure and easily be able to replicate the project.