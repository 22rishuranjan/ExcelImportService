```
# Excel Import Service (Simple Overview)

This project provides a simple flow to:

- Download an Excel template  
- Upload a filled Excel file  
- Validate every row using attributes  
- Save the original uploaded file and validation result in MongoDB  
- Download the same Excel back with a new "Status" column showing all errors per row  
- Access the same functionality through REST APIs and GraphQL

The goal is to give users a clean experience when importing structured data into the system.

---

## Excel Template Flow

### Template endpoint



##### GET /api/countries/template


What this returns:

- An Excel file with columns  
  - Code  
  - Name  
  - IsActive  
  - StartDate  
- Built-in Excel validations  
- Header comments describing allowed values  
- Sample row  

Users fill this file and re-upload it.

---

## Upload Flow

### Upload endpoint



##### POST /api/countries/upload
Form field: file (multipart/form-data)


What happens internally:

1. File stream is copied into memory and stored as raw bytes.
2. A new `ImportJob` record is created in MongoDB with
   - FileName  
   - StartedAtUtc  
   - Status = Running  
   - OriginalFile = uploaded file bytes  
3. Each row is validated using attribute-based rules on the model.
4. If a row has errors, an entry is added to `ImportJob.Errors`.
5. If no errors for a row, the parsed entity is inserted into the database.
6. Job is updated with:
   - TotalRows  
   - SuccessCount  
   - FailureCount  
   - CompletedAtUtc  
   - Status (Completed or Failed)

Response contains:

- ImportJob id  
- Upload summary  
- Error list  

You use this `id` to download the error-annotated Excel later.

---

## Download Annotated Excel (Error Report)

### Error report endpoint



#### GET /api/countries/import/{id}


What this does:

- Loads ImportJob from MongoDB  
- Reads `OriginalFile` (the uploaded Excel)  
- Opens the file using ClosedXML  
- Appends a new column named `Status`  
- Combines all errors for that row into a single text cell  
- Returns the modified Excel file for download  

This gives users their own uploaded data, but with validation errors highlighted clearly.

---

## GraphQL Support

The same functionality is exposed via GraphQL:

### Download template (base64)

```graphql
query {
  countriesTemplate {
    fileName
    contentType
    contentBase64
  }
}


### Upload Excel file

```graphql
mutation ($file: Upload!) {
  uploadCountriesExcel(file: $file) {
    job {
      id
      fileName
      status
      totalRows
      successCount
      failureCount
      errors {
        row
        column
        message
      }
    }
  }
}


### Get job details

```graphql
query ($id: String!) {
  importJob(id: $id) {
    id
    fileName
    status
    totalRows
    successCount
    failureCount
    errors {
      row
      column
      message
    }
  }
}


## Internal Processing (Short Explanation)

* TemplateService

  * Builds the Excel file users download
  * Adds comments and Excel validation rules

* ImportService

  * Reads Excel
  * Validates each row using attributes
  * Saves ImportJob + Errors
  * Saves original Excel file bytes

* ErrorReportService

  * Reopens the original file
  * Adds a "Status" column
  * Writes all error messages for each row

This keeps the system simple while giving users a clear import-and-fix workflow.

---

## Summary

Users can:

* Download a correct Excel format
* Fill and upload it
* See detailed errors
* Download the same file enhanced with error messages
* Fix it and upload again

Developers can:

* Add new import types easily by creating new models with attributes
* Reuse the same validator, error generator, and workflow

```
```
