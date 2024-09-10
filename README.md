# HybridOnBoarding.ps1

## Overview

> **Script Overview**
> 
> This script creates new users in Active Directory and Exchange Online based on a CSV file. It additionally copies group memberships and manager attributes from a template user.
> 
> **Features:**
> - Connects to Active Directory and Exchange Online.
> - Creates new AD user accounts and Exchange mailboxes.
> - Sets various user attributes such as description, display name, email address, and office.
> - Copies group memberships and manager attributes from a template user.
> - Sets a default password for the new user accounts.
> 
> **Usage:**
> 1. Prepare the CSV file.
> 2. Modify the script to match your environment.
> 3. Run the script.

## Prerequisites

- Active Directory Module
- Exchange Online Management Module

## Parameters

- **CSVPath**: Path to the CSV file containing user information.

## Example CSV File

```csv
FirstName,LastName,Type,Title,Fullname,EmployeeID,LOCATION
John,Doe,Front Office,Front Office Clerk,John Doe,12345,ER
Jane,Smith,Administrator,Manager,Jane Smith,67890,ADMIN OFFICE

