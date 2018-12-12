# stickyants-sp-util 

A set of utility functions and classes to support SharePoint development using JSOM. 
The library simplifies some of the tasks and uses ES6 promises to wrap `executeQueryAsync` calls
such that they work nicely together. 

## 1 Term Store (Taxonomy)

### 1.1 Working with the term store 

```javascript 

import {createTermStore} from 'stickyants-sp-util'; 


var store = createTermStore(); 

//to get all terms of a term set
var terms = await store.getAllTermsByTermSetId('11365da4-d7cd-43b3-9d28-9d43b0313f2f'); 

```

## 2 Link Generation


### 2.1 Generate link to user picture by user email address

```javascript 

import {getUserPictureLink} from 'stickyants-sp-util'; 


var pictureUrl = getUserPictureLink('someuser@sometenant.com');

```

### 2.2 Generate delve profile of user by email user email address

```javascript 

import {getUserDelveLink} from 'stickyants-sp-util'; 


var delveLink = getUserDelveLink('someuser@sometenant.com');

```

### 2.3 Generate a link to specific version of the current page using UI Version 

```javascript 

import {createVersionLink} from 'stickyants-sp-util'; 

//first major version link 
var delveLink = createVersionLink(512);

```


## 3 Using async/await to work with JSOM 

```javascript 

import {executeOnContext} from 'stickyants-sp-util'; 


var ctx = new SP.ClientContext();
var web = ctx.get_web(); 
ctx.load(web); 
await executeOnContext(ctx); 
console.log(web.get_title()); 

```

## 4 Working with folders

### 4.1 Creating a nested folder structure in SharePoint list 

```javascript 

import {createFolderInListIfNotExist} from 'stickyants-sp-util'; 

var ctx = new SP.ClientContext();
var web = ctx.get_web(); 
var pages = web.get_lists().getByTitle('Pages'); 
var folder = await createFolderInListIfNotExist(ctx,pages,'folder1/folder2/folder3');//returns folder3

```

### 4.2 Check if folder exists in List 

```javascript 

import {folderExistsInList} from 'stickyants-sp-util'; 

var ctx = new SP.ClientContext();
var web = ctx.get_web(); 
var pages = web.get_lists().getByTitle('Pages'); 
var result = await folderExistsInList(ctx,pages,'folder3'); //returns true/false

```

## 5 Publishing Infrastructure 

### 5.1 Creating a publishing page

```javascript 

import {createPublishingPage} from 'stickyants-sp-util'; 

var ctx = new SP.ClientContext();
var result = await createPublishingPage(ctx,'somfile','My Content Type Name'); 

```

### 5.2 Getting page layout that is associated with a specific content type 

```javascript 

import {getPageLayoutItemByAssociatedContentType} from 'stickyants-sp-util'; 

var ctx = new SP.ClientContext();
var result = await getPageLayoutItemByAssociatedContentType(ctx,'My Content Type Name'); 

```

## 6 SharePoint Error Codes (JSOM)

```javascript 

/**
 * JSOM error codes, the possible values of error codes that JSOM 
 * can return 
 */
export enum SPErrorCode {
    /**
     * The accessed item has been deleted
     */
    ItemDoesNotExist = -2147024809,
    /**
     * Some unknown error
     */
    GenericError = -1,
    /**
     * User does not have permission to access/perform operation 
     */
    AccessDenied = -2147024891,
    /**
     * A document with the same name already exist and the user is attempting to add new file/folder
     */
    DocAlreadyExists = -2130575257,
    /**
     * A version that is more recent has been saved 
     */
    VersionConflict = -2130575339,
    /**
     * List item is in recycle bin 
     */
    ListItemDeleted = -2130575338,
    /**
     * A field value that has been provided is invalid
     */
    InvalidFieldValue = -2130575155,
    /**
     * Operation not supported
     */
    NotSupported = -2147024846,
    /**
     * 
     */
    Redirect = -2130575152,
    /**
     * 
     */
    NotSupportedRequestVersion = -2130575151,
    /**
     * A specific field validation has failed 
     */
    FieldValueFailedValidation = -2130575163,
    /**
     * Item update failed validation rules
     */
    ItemValueFailedValidation = -2130575162,
}


```