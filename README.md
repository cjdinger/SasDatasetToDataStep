# SAS custom task example: SAS Data Set->DATA Step
***
This repository contains one of a series of examples that accompany
_Custom Tasks for SAS Enterprise Guide using Microsoft .NET_ 
by [Chris Hemedinger](http://support.sas.com/hemedinger).

This particular example goes with
**Chapter 12: Abracadabra: Turn Your Data into a SAS Program**.  It was built using C# 
with Microsoft Visual Studio 2010.  It should run in SAS Enterprise Guide 4.3 and later.

## About this example
This task example demonstrates several techniques that you can apply in your own tasks, including:

- Use ADO.NET to connect to and read SAS data. (ADO.NET is the method for working with data sources in .NET programs.)
- Use SAS DICTIONARY tables to discover data set and column attributes for the source data. Then, use that information to influence task behavior.
- Use the SASTextEditorCtl control from SAS.Tasks.Toolkit.Controls to preview the SAS program in the SAS color-coded Program Editor.
- Encapsulate the meat of the task -- the business logic that reads data and creates a SAS program -- into a separate .NET class. This makes the task more maintainable and enables you to reuse the business logic in other contexts.
- Use a special task interface, named ISASTaskExecution, to implement the work of the task. This enables SAS Enterprise Guide to delegate all task processing to your task so that it can perform work that can't be done easily in a SAS program.



