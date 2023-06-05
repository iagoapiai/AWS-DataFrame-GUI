# StatusBolt
Python script that accesses aws s3!

- Enters a bucket;
- Takes all the names of the files inside the bucket;
- Assembles a list and compares it with the current date;
- If the current date is the same as the name of the file "horario" it shows online;
- Shows the filtered list with online and offline asset management in a pysimplegui graphical interface;
- Create an Excel file in the main folder.

OBS: This code will only work if you redo the modifications of the dataframe that is taken from the cloud, it also needs to contain a folder called 'statusbolt' to save the excel file and an excel file inside that folder called 'idbolt' for it to replace the names with ids. 

Don't forget to put your aws credentials




![Gui](https://github.com/iagoapiai/AWS-DataFrame-GUI/assets/116030785/b994d2d8-d818-4f90-8d72-9643fae47b28)





