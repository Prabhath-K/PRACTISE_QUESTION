Handling Extreme Sales Values in Excel

If management insists on using the Average to represent sales, the impact of extreme values can be reduced by using Excel's TRIMMEAN function instead of a normal AVERAGE.

Excel Formula

=TRIMMEAN(FILTER(C2:C11,(D2:D11="South")*(C2:C11<>"")),0.2)

Explanation

FILTER(C2:C11,...) filters the sales values.

D2:D11="South" includes only records from the South region.

C2:C11<>"" ignores blank sales values.

TRIMMEAN(...,0.2) calculates a trimmed average by reducing the influence of extreme values.

0.2 represents a 20% trim from the extreme ends of the dataset.

Why Use TRIMMEAN?

A normal average can be heavily affected by unusually large or small values.

For example, the dataset contains an extreme sales value of:

₹300,000

This value can significantly increase the normal average and make it less representative of typical daily sales.

Using TRIMMEAN provides a more reliable average while keeping all original rows in the dataset.

Normal Average Formula

For comparison, the regular average formula would be:

=AVERAGE(FILTER(C2:C11,(D2:D11="South")*(C2:C11<>"")))

However, this calculation is more sensitive to extreme values.

Conclusion

TRIMMEAN is a better choice when management requires an average but the dataset contains outliers or extreme values.
