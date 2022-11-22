# Pizza_Reports_Excel
Based on pizza orders, determining the ingredients Pizza Maven should buy in order to become
a more efficient restaurant in terms of stock management, saving the results of such suggestions
to an xml file and building some reports in excel format with key stats for Pizza Maven´s COO.
To do so, in the first place you have to run in the python terminal the following command:
"pip install -r requirements.txt", which will automatically download the necessary libraries for
the programs to work correctly. In this repo you will find two executable python files:
"maven_pizzas_2016_excel.py" and "crear_excel.py". The first one uses the files given by Maven Pizzas
"order_details.csv", "orders.csv", "pizzas.csv" and "pizza_types.csv" to predict and estimate the
ingredients Pizza Maven´s manager should by each week of the following year. Additionaly, this
program generates an xml file ("informe_pizzas_maven_2016.xml") which corresponds to a brief report
and analysis of the data's quality. The other program, "crear_pdf.py", will be the one in charge of
generating the excel file with the reports.

To build such prediction, the program "maven_pizzas_2016_excel.py" processes the "order_details.csv"
file, applying some indispensable changes to the pizza names and the quantity, so that every one of
them has the same structure to simplify the next operations and treatment of the dataframes. Then,
the program sums up all the pizzas that were ordered last year, differentiating pizza type and size.
Then, divides that total quantity by 52 (number of weeks in a year) and takes the roof of the result
of the division (for example, 300/52 = 5,769 => 6). Once this procediture has been done for every kind
of pizza, we have the number of pizzas of each type ordered in an average week. Now, the program
multiplies each pizza´s ingredientes by the number of pizzas of such type ordered in one week by another
factor, related with the size. For each size, we consider the following units of each ingredient in one
particular pizza: {s:1, m:2, l:3, xl: 4, xxl: 5}

Once the number of ingredients has been calculated, leaving room for some misestimation, an xml file is
generated with the name "compra_semanal_ingredientes.xml". This file is a kind of shopping list Pizza
Maven´s manager should follow each week to be efficient and throw almost no food at all. The same
information is also uploaded to a csv named "compra_semanal_ingredientes.csv", which will be used, together
with "pizzas.csv", by the program "crear_excel.py" to build the necessary sheets with data and charts that
will be included in the excel file "Pizza_Reports.xlsx".
