# Event Handler: Display Account Primary Contact Details On Case Form

## Description

This project contains a simple script intended for use with a Power Platform Dataverse based App using Dynamics 365 Customer Service. 
It is used on the 'Case for Interactive Experience' form, populates a newly added quick view form (must be added by user) with information from the Primary Contact belonging to the Cases Customer if the Customer is an Account.
It also makes it mandatory to choose a new Contact for the Account if no Primary Contact currently exists.

## Behaviour

When the form loads or the Customer control is changed, the script runs to check if the Customer is an Account. 
If so, it then checks if the Account has a Primary Contact. If it does, the Primary Contact will be set as the Contact of the Case, and the quick view form will show it's details.
If the Account had no Primary Contact, the attribute for the Cases Contact will be made mandatory, and a lookup control for selecting a Contact will be made visible.

When the form saves, the script checks whether the Cases Contact attribute has been made mandatory.
If so, the Contact will be set as the Primary Contact of the Account.
If the Contact chosen was not already associated with the Account, the user will be shown an error, as a Contact must be associated with an Account before it can be set as its Primary Contact.

## Function Description
### showAccountPrimaryContact

Function intended to be called on form load and Customer control change. Calls other functions and branches depending on if Account has Primary Contact. 
Also catches errors thrown by other functions and creates an error dialogue with their message. If an error has no message, no dialogue is made as this is used by the script to end early.

### retrieveContact

Retrieves the Accounts Primary Contact using the Xrm Web Api. Throws an error if any issue occurred while retrieving the record from the Dataverse. 
Also ends the script early if the Customer assigned to the case is a Contact, which ends the script without any error message to the user, as this is an acceptable circumstance.

### getCustomerAttribute

Gets the "customerid" attribute from the Case record. Throws an error if the attribute is null, as this suggests that the Case table does not have a "customerid" column in the current environment.

### getCustomerRecord

Gets the record details of the Customer from the Customer attribute. Ends the script early if Customer had no value assigned, as this is normal when creating a new case.

### populateForm

Sets the value of the "primarycontactid" attribute of the Case to the primary Contact of the assigned Account. 

### hideEmptyFields

Conditionally hides the fields of the Primary Contact quick view form if their associated data is empty. Throws an error if the quick view form could not be found in the main form.

### toggleContactMandatory

Makes the "primarycontactid" Case column mandatory and shows the corresponding control on the form, or does the opposite depending on the boolean parameter.

### saveContactIfRequired

Function intended to be called on form save. Calls other functions as well as conditionally ending the script early if the "primarycontactid" column is not mandatory. 
Also catches all errors and creates an error dialogue with the errors message.

### checkAccountOwnsContact

Throws an error and cancels saving of the form if the given Contact is not associated with the Cases Account.

### contactHasAccountAsParent

Returns true if the "parentcustomerid" value of the Contact associated with the supploed Contact Id matches the supplied Account Id, and false otherwise.
Also throws an error if there was an issue retrieving with the web api.

### saveAccountPrimaryContact

Updates the supplied account's Primary Contact to be the supplied Contact. Also throws an error if there was an issue updating with the web api.
