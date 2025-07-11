## Program Shortcomings and potential solutions

**Finding The Right Vendor Name**

Currently, the program splits the file path into multiple strings using specific delimiters as separators.
It then removes the file extension (e.g., .ect). This process truncates the strings, leaving only the "base," which is the actual file name.
Next, the matching algorithm uses the remaining portion to find a match in the Excel spreadsheet that is at least 60% accurate.
Additionally, there are certain exclusions added to an 'ignore' list to improve the accuracy of the matching.

Before executing the matching algorithm, the program uses regular expressions (re) to identify the four-digit IRD.
There are inclusions for an apostrophe and a letter, as the IRDs in the Excel file sometimes contain unusual apostrophes, and the file names are not consistent.
If the program finds a match for the IRD, it uses the corresponding vendor name; if not, it resorts to guessing the vendor name.

The main issue is that any exclusions not on the ignore list can cause the matching algorithm to fall below the minimum 60% accuracy threshold.
While I can keep adding more keywords to improve the matching process, this approach is not efficient due to inconsistencies.
Moreover, it won't help in cases where the file name does not include the vendor name at all and only consists of the file type.

**PDF Specific: Dating**

The program is precise but lacks accuracy. I can enhance the accuracy by simplifying the target.
Specific rules can be established for the file type.
These rules would likely eliminate the need for validation (except for clearly incorrect dates), and the parsing process (which transforms date strings into datetime objects) could be streamlined within the ***extract_date_strings*** function.
With that, I could consider removing the ***find_effective_date_context*** function entirely.
But similar to adding exclusions for the vendor name guessing, it is not dynamic.

A more advanced approach would be to model the process of finding dates.
This would be preferable in cases where documents do not consistently follow dating practices, even if they belong to a category for which we have defined rules.
Given the file type and context, a weighted model could more accurately identify the correct date.
If trained properly, this model would result in a higher degree of both accuracy and precision. But my programming skills aren't quite there yet lol.
