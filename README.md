# djb-brevet-card-generator
A Google Sheets script used to generate brevet cards from a spreadsheet of registered user information

Many assumptions are currently made:
- Document in which this script will be run has a sheet called Submissions, which contains the rider-specific information
- Document in which this script will be run has a sheet called Controles, in a specific format, which contains route-specific information
-- Should have 4 columns (Controle Name, Controle Address, Controle Times, Controle Distance)
-- Should have 11 rows. First is header, 11th has "Bike Inspection" in the first (Controle Name) columns
-- Controle Times values should be in the format "open: 03/08 10:19 close: 03/08 12:00"
-- Controle Distance values should be in the format "Mile 28"
- Template files for the inside and outside of the brevet card exist, and are correctly specified by id in the top of the script.
- We wish to generate one blank card for every registered rider.
- We won't have more than 9 controles
- We'll print all the fronts first, then manually flip the papers and print all of the backs.

Missing features: 
 - Generation of the backs of the cards is hard-coded to 16 cards.
 - I'd like to add the ability to intermingle fronts and backs, e.g. page 1 of the generated doc is two fronts, page 2 is two backs, page 3 is two fronts, etc.
 - I'd like to add the ability to specify the number of desired blank cards
 - I'd like to add a visual way to set the values for each generation run (like output file name, start time, organization, ride type, etc.)


