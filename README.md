Instructions for Physical Count

1) Export products to items.csv using the following column structure

    itemNumber,name,unitOfMeasure,binLocation,onHand,price

2) Run `fix_csv.rb items.csv` to ensure proper quoting.

3) Run `generate.rb` to create worksheets to be printed, and files in data_entry to be modified after the count

4) Run `tally.rb` to generate worksheets to be printed for rechecks

5) After completing data entry and recheck move worksheet to done folder

6) After all data entry worksheets are complete, run `final_variance.rb` to generate final_vriance.xls file

7) Run `export.rb` to generate export files for ERP-ONE

8) Open WACNT in ERP-ONE, select all products

9) import files from export, confirming values in data entry

10) print and save variance report

11) update physical count in WACNT, print final report  


