<#
	Replace-CityName.ps1
	Created By - Kristopher Roy
	Created On - 27 Jul 2023
	Modified On - 29 Jul 2023

	This Script takes the export csv from wordpress and allows you to replace names
#>

#This is a function to handle multiple name cities and set the first letters to capitol
function capitalize_city($city) {
  $words = $city.split(" ")
  $capitalized_words = @()
  foreach ($word in $words) {
    $first_letter = $word.Substring(0, 1).ToUpper()
    $rest_of_word = $word.Substring(1).ToLower()
    $capitalized_word = $first_letter + $rest_of_word
    $capitalized_words += $capitalized_word
  }
  $capitalized_city = $capitalized_words -join " "
  return $capitalized_city
}

#This is the csv file to import for conversion replace the folder and file path of this file
#The headers in this file must be "Service Name, Slug, Heading, SubHeading, Content 1, CTA Heading, CTA SubHeading, Content 2" If these headers change, then the script will need to be updated
$csvwebfile = import-csv 'C:\Projects\Kova\Floor Care Installation - Hardwood Floor Refinishing Cities (4).csv'

#This is the csv file with the list of cities to place in the file replace the folder and file path of this file
#The header in this file must be "Service Name" 
$NewCityList = import-csv 'C:\Projects\Kova\city names - Floor Care Installation - Hardwood Floor Refinishing Cities (5).csv'

#loop through the citylist and create new csvwebfiles to import back into wordpress
Foreach($city in $NewCityList)
{
    $oldcity = $csvwebfile.'Service Name'
    $oldcitylower = $oldcity.ToLower()
    $newcitylower = ($city.'Service Name').ToLower()
    $newcityforurl = $newcitylower -replace '\s',''
    $newcityUpper = capitalize_city $newcitylower
    $csvwebfile.'Service Name' = $newcityUpper
    $csvwebfile.Slug = $newcitylower+"-ca"
    $csvwebfile.Heading = $csvwebfile.heading.Replace($oldcity,$newcityUpper)
    $csvwebfile.Subheading = $csvwebfile.subheading.Replace($oldcity,$newcityUpper)
    $csvwebfile.'Content 1' = $csvwebfile.'Content 1'.Replace($oldcity,$newcityUpper)
    $oldcityrmvspace = $oldcitylower -replace '\s',''
    $regex = "(?i)\b$oldcityrmvspace-ca\b"
    $csvwebfile.'Content 1' = $csvwebfile.'Content 1' -replace $regex, "$newcityforurl-ca"
    $csvwebfile.'Content 1' = $csvwebfile.'Content 1'.Replace($oldcity+"-ca",$newcitylower+"-ca")
    $csvwebfile.'Content 1' = $csvwebfile.'Content 1' -creplace $oldcitylower,$newcitylower
    $csvwebfile.'Content 1' = $csvwebfile.'Content 1' -creplace $oldcity,$newcityUpper
    $csvwebfile.'CTA Heading' = $csvwebfile.'CTA Heading'.Replace($oldcity,$newcityUpper)
    $csvwebfile.'CTA SubHeading' = $csvwebfile.'CTA SubHeading'.Replace($oldcity,$newcityUpper)
    $csvwebfile.'Content 2' = $csvwebfile.'Content 2' -replace $regex, "$newcityforurl-ca"
    $csvwebfile.'Content 2' = $csvwebfile.'Content 2'.Replace($oldcity+"-ca",$newcitylower+"-ca")
    $csvwebfile.'Content 2' = $csvwebfile.'Content 2' -creplace $oldcitylower,$newcitylower
    $csvwebfile.'Content 2' = $csvwebfile.'Content 2' -creplace $oldcity,$newcityUpper
    
    #Exports the content into a new individul csv files named whatever the new city is -export.csv
    #$csvwebfile|export-csv C:\Projects\Kova\$newcitylower-export.csv -NoTypeInformation -Encoding Unicode

    #Exports the content into a new individul csv file named newcities-export.csv
    $csvwebfile|export-csv "C:\Projects\Kova\newcities-export.csv" -NoTypeInformation -Encoding Unicode -append
}