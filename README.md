# Stanford Transfusion Service and Stanford Blood Center Inventory and Order Management

Inventory management is both a major challenge and an integral part of hospital transfusion service (HTS) and blood centers (BC) operations. At our HTS, the general inventory accounts for over 50 product categories broken down by component, blood type, irradiated status, and CMV-serology status. Products are stored in Safetrace TX Database by unit number including ISBT code, this data is extracted through Crystal reports. This program tabulates the number of units in stock for platelets, plasma, and cryoprecipitate using categorization via ISBT code and other attributes. For RBCs, the program uses a linear regression model specific to Stanford's HTS to infer the levels because the raw electronic inventory report comprised both the general inventory and physically sequestered units (e.g. special antigen units, preemie units, and cross-matched units). The macro then compares the calculated inventory to the desired "Par" (aka buffer/stock) levels and outputs the suggested shipment. There is a final adjustment step where surplus CMV-negative units are counted toward CMV-positive/untested units as of v3.1+.

## Prerequisites

Excel with macros enabled.

CSV file from Crystal reports. The file begins the following headers:

```
Received	Type	DIN	P.Code	Donor	Division	ABO	Rh	CMV	UIP	Exp. Date	Exp. Time	Status
```

The CSV is generated with the following SQL query:


> SELECT "PRODUCT_INVENTORY"."UNIT_NO", "PRODUCT_INVENTORY"."STANDARD_PRODUCT_CODE", "PRODUCT_INVENTORY"."PRODUCT_ID", "PRODUCT_INVENTORY_ACTIVITY"."TRANSACTION_CD", "PRODUCT_INVENTORY"."PRODINV_ID", "PRODUCT_INVENTORY_ACTIVITY"."ENTERED_DATETIME", "PRODUCT_INVENTORY"."DONATION_TYPE_CD", "PRODUCT_INVENTORY"."DIVISION", "PRODUCT_INVENTORY"."INVENTORY_STATUS_CD", "PRODUCT_INVENTORY"."ABO_CD", "PRODUCT_INVENTORY"."RH_CD", "PRODUCT_INVENTORY"."UNITS_IN_POOL", "PRODUCT_INVENTORY"."LOCATION_ID", "PRODUCT_INVENTORY"."EXPIRATION_DATE", "PRODUCT_INVENTORY"."EXPIRATION_TIME"
> FROM   "EBIS"."PRODUCT_INVENTORY" "PRODUCT_INVENTORY" INNER JOIN "EBIS"."PRODUCT_INVENTORY_ACTIVITY" "PRODUCT_INVENTORY_ACTIVITY" ON "PRODUCT_INVENTORY"."PRODINV_ID"="PRODUCT_INVENTORY_ACTIVITY"."PRODINV_ID"
> WHERE  "PRODUCT_INVENTORY_ACTIVITY"."TRANSACTION_CD"='RD' AND ("PRODUCT_INVENTORY"."INVENTORY_STATUS_CD"='A' OR "PRODUCT_INVENTORY"."INVENTORY_STATUS_CD"='G' OR "PRODUCT_INVENTORY"."INVENTORY_STATUS_CD"='U') AND "PRODUCT_INVENTORY"."LOCATION_ID"='TS'
> ORDER BY "PRODUCT_INVENTORY"."PRODINV_ID"


Additionally, the program requires the date and time in the CSV filename of when the CSV was generated. This requires the format “ … MM.DD.YYYY … HHmm … .csv” in the filename where: 

1.	“MM.DD.YYYY” is the month, date, and year separated by periods 

2.	“HHmm” is the hours and minutes in military time with no separation.

3.	“…” is any length or number of characters not matching 1 and 2.

4.	The file is csv filetype

The process we are using only looks for the date and time in the filename and ignores the other information. Any other parts of the filename are for user clarity.

```
TSI MM.DD.YYYY at HHmm.csv
```

```
TSI 05.30.2017 at 0730.csv
```

## Operating Protocol

### Generate Shipping Projection

1. Open the project file e.g. “TS RBC inventory projection vX.X.xlsm”

2. Enable macros in Excel.

3. Click “Generate Shipment Prediction” button in the Shipping_Projection worksheet.

4. A dialog box will prompt the user to select the CSV file.

5. The Shipment Prediction for each product will appear in the Shipping_Projection worksheet. Please check that the date and time in cell L1 match the CSV file and the current date and time for the order being made.

6. Click the “Save copy” button in the Shipping_Projection worksheet. This will save a copy of the current state in the same folder.

7. The supervisor will print the Shipping_Projection worksheet and fullfill the order. A copy will be faxed to the HTS.

8. If SBC does not have enough inventory to fill the order for a product, the SBC tech will check the hand count and determine if they can fill the order according to the hand count. The SBC tech will write the number of the actual number shipped in the “Actual Shipment” column, if the number send in the order deviates from the number in the “Shipment Prediction” column. A comment should explain the reason for deviation, i.e. “insufficient inventory.” The printed sheet will be signed and dated as a record of the order.

9. If the application fails, please use the hand counts and fill to par levels as in the past. 

### Updating Par Levels (and offsets) and Linear Regression Equations

The Dashboard worksheet contains columns "Equation", "Offset Adjust", and "Restock Default."

1. Unprotect the sheet using "shcsbc" password

2. Change the desired value

3. Reprotect the sheet with "shcsbc" password

Note: it is strongly discouraged to change any of the other cells in this sheet as the program extensively references "Product Type" and "Blood Type" columns to function. Excercise caution when changing equations as the syntax is extremely rigid for parsing the equations.

```
y = 0.6466x - 10.935
```

### Updating ISBT codes

Most of the common ISBT codes and product types in use at Stanford HTS are entered, but this is currently not an exhaustive list of all possible ISBT codes. The program makes some assumptions if unknown ISBT codes are encountered to keep the experience smooth for the end-user. One of the assumptions is that the number of these units is small so it will gloss over these since one or two should have neglible effect. In monitoring this system for several months, we've found that new ISBT codes are rarely encountered and we have added them to the list each time. However, there could be a larger discrepancy If new codes are added that are used often. There are exhaustive lists available, but annotating hundreds of codes that are unlikely to be ncountered and also adding them might clutter the existing list.

The ISBT key is found on the "Product Codes List" worksheet.

The important columns referenced by the program are "product_code",	"product", "irradiated", and "Inventory Category". The "attribute" category is only referenced for RBC units to look for the "dry" attribute.

## Built With

* Microsoft Visual Basic for Applications 

## Authors

* **Hamilton Tsang** - *Initial work* - [Hamiltologic](https://github.com/Hamiltologic)

See also the list of [contributors](https://github.com/placeholderplaceholder) who participated in this project.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments


