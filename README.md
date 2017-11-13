## SAP BO Analysis for Office (BOAO) Automation

This tool allows you to automate refresh of workbooks with BO Analysis data sources.
Tool helps to automate change of Variables (Prompts) and dimension Filters (Background Filters), and then refresh process itself.
On top of this, you can configure additional actions like "Save As", "Save As & Email", "Refresh All", Run another specific macro etc.

Solution consists of only one worksheet (VBA code is inside it)

- [BOA Control Panel.xlsb](https://github.com/IvanBond/SAP-BOA-Automation/blob/master/BOA%20Control%20Panel.xlsb)

Worksheet can be easily moved to your workbook using standard "Move worksheet" Excel action. Then just collect variables, set values and run Refresh.

# BOA Control Panel

Your future operational center. Control Panel is a worksheet, which includes
- tables defining scenarios of refresh and variables with their values
- VBA code

# How to use this tool

Assume you already have a workbook with BOA data sources and want to simplify refresh process. Let's call it 'Target Workbook'.

0. Open 'Target Workbook' and 'BOA Control Panel' side by side in one Excel application.

1. Move worksheet 'Control Panel' to 'Target Workbook'

2. Press 'Collect Variables'. Macro will make inventory of data sources and their prompts / variables / filters.
    
3. Specify necessary settings, such as "Scope", "Refresh?", "Order", values for Variables and Filters.
Use formulas to make values of your variables dynamic, then you no longer need to change them manually.

You are ready to refresh!

# Optional steps

- If you don't want to enter your password each time - follow the instruction in comment for 'Path to file with passwords' cell.

- You can specify macros that should be executed before BOA refresh and after (e.g. for your saving/mailing scenario).

# What is 'Scope'?

Scope defines set of settings for data sources and sets of variables.

Assume you want to refresh same workbook for two different Sales Organizations. 

Easy. Just define two Scopes with corresponding values for variables.

Using Scopes you may define very advanced scenarios of refresh.

E.g. imagine report when you need to execute 10 queries for current and previous year. Without Scopes it would be 20 queries, 10 for each year. But with Scope you may leave only 10 queries.
Define two Scopes - Prior Year, Current Year. Enable 'Refresh All Scopes' option. Using formulas for variables, force them to calculate corresponding to active Scope values. Add simple macro that will copy data after 'Prior Year" scope refresh is done to another worksheet. Use it in 'Macros After' for the last data source of PY scope. Then after refresh of all scopes you will have static data of PY on one worksheet and data sources with CY on another.

If you run refresh from outside of workbook, e.g. like it is shown in [Sample Refresher VB script](https://github.com/IvanBond/SAP-BOA-Automation/blob/master/Refresher%20Sample.vbs) - you can even run refresh in parallel.
