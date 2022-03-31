# price-manager

Application User Intarface 
  - ex.: Excel (stored on local computer)
  - 
Datamanagement System
  - ex.: Access (stored on sharepoint)
  - UserManagementTable (Ex.: Client, Approver, Manager, Developer)

# Workflow

![Workflow diagram](https://user-images.githubusercontent.com/25910991/161052023-c44f42b5-893e-4f9e-96ae-2490e50d10a1.png)

# I. Client opens application:
	
	Completes data
	Saves data (data to database)
	Auto Notifie Approver 


# II. Approver opens application:
		
	List of all prices (Data from database)
	Filters 
	Approve (marks as aproved) 
	Decline (marks as unaproved, notifies responsable user)
	Export filtered data to new excel workbook
    
    
# Application Form Context:
## Fields available for client
	Customer field
		- field features: 	-> required
					-> validation (maxLength 6, starts with nr. 4 or 5)
	Material field
		- field features: 	-> required
					-> validation (maxLength 8, starts with nr. 5)
	Price field
		- field features: 	-> required
					-> validation (maxLength 6, 4 numbers + 2 decimals)
	Currency field
		- field features: 	-> required
					-> dropdown static list (EUR, USD, GBP, PLN)
	Price Unit field
		- field features: 	-> required
					-> validation (maxLength 4, no decimals)
	Unit of Measure field
		- field features: 	-> required
					-> dropdown static list (KAR, RO, ST, KG, LM, M2)
	Valid from field
		- field features: 	-> autocompleted todays date
		
	Valid to field
		- field features: 	-> precompleted date 31.12.9999
		
	Add Button
	Edit Button
	List (Currently added prices)

## Fields available for approver:
	Customer Filter 
	Approved/declined Filter
	Approved/declined Date Filter
	Saved/notSaved Date Filter
	Export
