# Plugin: Cancel Case Creation If Customer Has Existing Active Case
## Description

This project contains a simple plugin intended for use with a Power Platform Dataverse based App using Dynamics 365 Customer Service.
It cancels the creation of any Case where its Customer already has an active Case associated with it.

## Behaviour

Any time a Case is created, the connected Dataverse is checked to see if any existing case exists with the same Account Id as its Customer.
If so, the new Case is cancelled and an error displayed to the user, otherwise no action is taken.

## Function Description

### Execute

This is the main function executed when the plugin is run. It throws an InvalidPluginExecutionException if an existing Case with the same Customer is found.
It also catches all errors, and rethrows them along with creating a trace log. It also converts all thrown exceptions into InvalidPluginExecutionExceptions.

### GetEntity

Retrieves the new record entity from the input parameters. Throws an exception if the Entity is not found in the input parameters or the Customer Id is not found in the entity.

### BuildQuery

Returns a query to be used for retrieving Entities from the Dataverse. Top count is set to one as only one Case should ever be found at maximum.
A condition is added to the customerid to ensure it matches the customerid of the new Cases Customer, and another condition is added so that only
Cases with a status code less than 5 are included, as Cases with a status code 5 or higher are always either resolved or cancelled.

### CustomerHasExistingCases

Sends a request to retrieve the Entities which match the built query, returning true if the number of results is more than 0 and false otherwise.
