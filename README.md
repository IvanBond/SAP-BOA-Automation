## SAP BO Analysis for Offfice (BOAO) Automation

This tool allows you to automate refresh of your workbooks with BOA data soures.
If you have a workbook based on BOA and want to automate change of Variables or dimension Filters, and then refresh process itself - try this tool.

Solution consists of only one file

- [BOA Control Panel.xlsb](https://github.com/IvanBond/SAP-BOA-Automation/blob/master/BOA%20Control%20Panel.xlsb)

File contains only one worksheet, which can be easily moved to your workbook. Then just collect variables, set values and run Refresh.

# BOA Control Panel

Your future operational center. Workbook contains only one worksheet, which includes
- tables defining scenarios of refresh and variables with their values
- VBA code

# How to use this tool

Assume you already have a workbook with BOA data sources in it and want to simply refresh process. Let's call it 'Target Workbook'.

1. Prepare list of data sources and variables for your workbook. Press 'Collect Variables'.

    - Start empty copy of MS Excel.
    - Open 'Target Workbook' and 'BOA Control Panel'
    - Move 'ControlPanel' sheet to 'Target Workbook'
    - Press 'Collect Variables'
    
2. Specify necessary settings, such as "Scope", "Refresh?", "Order", values for Variables and Filters.

You are ready to refresh!

# Optional steps

- If you don't want to enter your password each time - follow the instruction in comment for 'Path to file with passwords' cell.

- You can specify macros that should be executed before BOA refresh and after (e.g. for your saving/mailing scenario).

# What is 'Scope'?

Somewhere I saw such definition: **Structured Computing Optimized for Parallel Execution**. Sounds neat!

In this particular solution, Scope defines set of settings for data sources and sets of variables.

Assume you want to refresh same workbook for two different Sales Organizations. 

Easy. Just define two Scopes with corresponding values for variables.

If you run refresh from outside of workbook, e.g. like it is shown in [Sample Refresher VB script](https://github.com/IvanBond/SAP-BOA-Automation/blob/master/Refresher%20Sample.vbs) - you can even run refresh in parallel.

# Tables on ControlPanel

![Scopes and Data Sources](https://bondarenkoivan.files.wordpress.com/2017/02/boa-automation-scopes-and-data-sources.png)

![Variables](https://bondarenkoivan.files.wordpress.com/2017/02/boa-automation-variables-table.png)
