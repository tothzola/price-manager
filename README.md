# price-manager

Application User Interface 
- ex.: Excel 
- 

Datamanagement System
- ex.: Access 
- Users Table (Client, Approver, Manager, Developer)

# Workflow
![Price Manager](https://github.com/tothzola/price-manager/blob/main/pictures/WorkFlow_PriceApprovalManager.jpg)

# I. Client opens application:
- Complets form data
- Saves data (data to database)
 -Auto Notify Approver 

# II. Approver opens application:
- List of all prices (Data from database)
- Filters 
- Approve (marks as aproved) 
- Decline (marks as unaproved, notifies responsable user)
- Export filtered data to new excel workbook   

# Application Form Context:	
## Client Fields (features and validation)

### Condition Type

**Feature:** entry is required, default value "Y0PR", prefilled.

### Sales Organization

**Feature:** entry is required, selection value "2961", prefilled.

### Distribution Channel

**Feature:** entry is required, selection values "01", "HD", "GY"

### Customer

**Feature:** entry is required, numeric filed (format: "######")

**Validation:** exact 6 char length, range should be between [399999] and [599999] 

**ex. for validation failed:** 
```sh
[strings] [any special characters] [23423] [23452345] 
```
>lenght is <> 6, contains invalid characters and not in range.

**ex. for validation passed:** 
```sh
[453123] [592314] 
```
>lenght is = 6, and in range.
---

### Material

**Feature:** entry is required, numeric field (format: "########")

**Validation:** exact 8 char length, range should be between [49999999] and [59999999] 

**ex. for validation failed:**
```sh
[strings] [any special characters] [43423] [43452345]
```
>lenght is <> 8, contains invalid characters and not in range.

**ex. for validation passed:** 
```sh
[51234567] [57654321]
```
>lenght is = 8, and in range.
---

### Price

**Feature:** entry is required, string field (format: "#.###,00") is currency

**Validation:** maximal 6 char length

**ex. for validation failed:**
```sh
[strings]  [any special characters other then "." or ","] [43452345]
```
>lenght is > 6 and contains strings

**ex. for validation passed:** 
```sh
[012] [512] [5123] [544321]
```
>lenght is = 6,  values displayed = [0,12] [5,12] [51,23] [5.443,21]
---

### Currency

**Feature:** selection is required, dropdownlist 
- list values: EUR, USD, GBP, PLN

### Price Unit

**Feature:** no entry is required, numeric field (format: "####"), prefilled value is "1"

**Validation:** maximal 4 char length

**ex. for validation failed:**
```sh
[strings] [any special characters] [43423] [43452345]
```
>lenght is > 4, contains invalid characters.

**ex. for validation passed:** 
```sh
[1] [12] [9999]
```
>lenght is < 4
---

### Unit of Measure

**Feature:** selection is required if Price Unit field hase a valid value, dropdownlist 
- list values: KAR, RO, ST, KG, LM, M2
---

### Valid from

**Feature:** entry is required, numeric field (format: "##.##.####") is date, autocompeted as todays date, user is allowed to change the entry.

**Validation:** exact 10 char length

**ex. for validation failed:**
```sh
[strings] [any special characters other then "."] [43423] [43452345]
```
>lenght is <> 10, contains invalid characters and it is not a date

**ex. for validation passed:**  
```sh
[10.02.2009] [10022009] [31.12.2022]
```
>lenght is = 10 and it is a valid date.
---

### Valid to

**Feature:** entry is required, numeric field (format: "##.##.####") is date, autocompeted as [31.12.9999] date, user is allowed to change the entry.

**Validation:** exact 10 char length

**ex. for validation failed:**
```sh
[strings] [any special characters other then "."] [43423] [43452345] 
[10.02.2009] <= [Valid from field entry]
```
>lenght is <> 10, contains invalid characters,it is not a date and equals or it is in passt date compared to Valid from field.

**ex. for validation passed:** 
```sh
[10.02.2009] [10022009] [31.12.9999]
```
>lenght is = 10, and it is a valid date.
---

## Buttons
- Add Button
- Edit Button
- List (Currently added prices)
- Save Button

## Fields available for approver
- Customer Filter 
- Approved/declined Filter
- Approved/declined Date Filter
- Saved/notSaved Date Filter
- Export Button
