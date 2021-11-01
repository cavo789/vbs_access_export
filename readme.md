# MS Access - Export code to flat files

![banner](./banner.svg)

> Export every code objects of a MS Access database (forms, macros, modules, queries and reports) to flat files, in a batch mode

## Table of Contents

- [Description](#description)
- [Install](#install)
- [Usage](#usage)
- [Sample](#sample)
- [Author](#author)
- [License](#license)

## Description

This VB script will export all code objects (forms, macros, modules, queries and reports) from a MS Access database / application (can be `.accdb` or `.mdb`) to flat files on your disk.

This way, you'll get a quick backup of your code and you'll be able to synchronize your code on a versioning platform like GitHub.

The script will start MS Access (hidden way), open the specified database, process every code object and export them, one by one, in a `\src\your_database.mdb` folder.

The `src` folder will be automatically created if needed and you'll find a sub-folder having the same name of your file (so you can have more than one exported file in the same src folder).

## Install

Just get a copy of the `.vbs` script, perhaps the `.cmd` too (for your easiness) and save them in the same folder of your database.

## Usage

Just edit the `.cmd` file and you'll see how it works: you just need to run the `.vbs` with one parameter, the name of your database.

## Sample

For instance, by running `cscript vbs_access_export.vbs C:\Christophe\db1.mdb` you'll get this:

```
Process database C:\Christophe\db1.mdb
Exporting sources to C:\Christophe\src\db1.mdb\

Export all forms
        Export form 1/2 - frmCustomers to Forms\frmCustomers.frm
        Export form 2/2 - frmInvoices to Forms\frmInvoices.frm

Export all macros
        Export macro 1/15 - mcrPrint to Macros\mcrPrint.txt
        [...]

Export all modules
        Export module 1/2 - Helper to Modules\Helper.bas
        Export module 2/2 - SQL_Server to Modules\SQL_Server.bas

Export all queries
        Export query 1/23 - qryCustomers to Queries\qryCustomers.sql
        Export query 2/23 - qryInvoices to Queries\qryInvoices.sql
        [...]

Export all reports
        Export report 1/8 - Invoice to Reports\Invoice.txt
        [...]
```

Once finished, you'll have a sub-folder called `src` with one file by object so, indirectly, you've a backup of your code :ok_hand:

## Author

Christophe Avonture

## License

[MIT](LICENSE)
