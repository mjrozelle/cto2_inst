*! cto2.ado - Stata module to import and minimally clean SurveyCTO data
*! Author: Michael Rozelle <michael.rozelle@wur.nl>
*! Version 0.0.2  Modified:  March 2023

// Drop the cto_read program if it already exists
cap program drop cto2_inst
// Define the cto_read program, which takes three arguments
program define cto2_inst, rclass
// instrument, then dataset, then dofile
syntax, ///
	INSTname(string) /// filepath to the Excel survey instrument
	TEXdirectory(string) /// filepath to the import dofile you want to create
	SURVEYNAME(string) // full name of survey

	
pause on

version 16

// Create a quiet environment to prevent unnecessary output
qui { 
	
preserve 
local original_frame `c(frame)'

cap confirm file "`dofile'"
if !_rc & "`replace'" == "" {
	
	display as error "file `macval(dofile)' already exists."
	display as error "add option {bf:replace} if you wish to overwrite it."
	exit 602
	
}

cap confirm file "`reshapefile'"
if !_rc & "`replace'" == "" & "`reshapefile'" != "" {
	
	display as error "file `macval(reshapefile)' already exists."
	display as error "add option {bf:replace} if you wish to overwrite it."
	exit 602
	
}

if "`identifiers'" == "" {
	
	local identifying_vars
	
}
else {
	
	local identifying_vars `identifiers'
	
}

if "`american'" == "" {
	
	local datestyle DMY
	
}
else {
	
	local datestyle MDY
	
}

if "`reshapefile'" != "" {
	
	local want_reshape = 1
	
}
else {
	
	local want_reshape = 0
	
}

*===============================================================================
* 	Import XLSforms
*===============================================================================

/* We're going to work in two frames - there's the "survey" sheet containing the 
questions, enablement conditions, and so on, as well as the "choices" sheet,
which gives us all the value labels. Rather than open and close a million datasets,
frames let us work on both these levels simultaneously.
*/

*===============================================================================
* 	The Data
*===============================================================================\

// Create a new frame called "qs" to hold the survey questions
tempname qs
frame create `qs' 
cwf `qs' 

// Import the "survey" sheet from the instrument Excel file specified by the "instname" variable
import excel "`instname'", firstrow clear sheet(survey) 

missings dropobs, force

if "`rename'" != "" local rencheck new_name

// Loop over a list of variables and ensure that they exist in the dataset
foreach v in type name calculation relevant repeat_count `rencheck' hint {
	cap confirm variable `v'
	cap tostring `v', replace
	if (_rc) {
		local missvars `missvars' `v'
		continue
	}
	else {
		local keepvars `keepvars' `v'
	}
}

cap confirm variable label 

if !_rc {
	
	rename label labelEnglishen
	clonevar labelStata = labelEnglishen
	
}
else {
	
	cap tostring labelStata, replace
	
}

// Keep only the variables that exist in the dataset
keep `keepvars' labelEnglishen labelStata

// Display a warning if any variables are missing
if "`missvars'" != "" {
	noisily display as result "You are possibly missing the variable(s): `missvars'!"
}

// Replace any dollar signs in the "labelEnglishen", "labelStata", "repeat_count", and "relevant" variables with pound signs
foreach v of varlist labelEnglishen labelStata repeat_count relevant {
	replace `v' = subinstr(`v', "$", "#", .)
}

// Replace any missing Stata labels with the English labels
replace labelStata = "" if labelStata=="."
replace labelStata = labelEnglishen if missing(labelStata)

replace type = strtrim(stritrim(type))
replace type = subinstr(type, " ", "_", .) ///
	if inlist(type, "begin group", "end group", "begin repeat", "end repeat")
	
// Remove any line breaks, dollar signs, and double quotes from the "labelEnglishen", "labelStata", and "relevant" variables
foreach var of varlist labelEnglishen labelStata relevant hint {
	replace `var' = subinstr(`var', char(10), "", .)
	replace `var' = subinstr(`var', "$", "#", .) 
	replace `var' = subinstr(`var', `"""', "", .)
	replace `var' = subinstr(`var', `"<b>"', "", .)
	replace `var' = subinstr(`var', `"</b>"', "", .)
	replace `var' = subinstr(`var', `"<em>"', "", .)
	replace `var' = subinstr(`var', `"</em>"', "", .)
	replace `var' = strtrim(stritrim(`var'))
	/* Stata will intepret dollar signs as globals, and as a result, we won't 
	be able to see the questions being referenced for constraints (or even 
	inside the question itself). By changing the dollar sign to a #, we can 
	still see these references in the variable labels and notes */
}

// Replace any full stops in variable names with an empty string
replace name = subinstr(name, ".", "", .) 

// Split the "type" variable into two variables, "type" and "type2"
split type 
// Note that "type2" will contain the value label name for categorical variables.

*------------------------------------------------------------------
*	Question Types
*------------------------------------------------------------------

gen preloaded = regexm(calculation, "^pulldata")
gen note = type == "note"

local numeric_formulae index area number round count count-if sum ///
	sum-if min min-if max max-if distance-between int abs duration
local regex_stubs: subinstr local numeric_formulae " " "|" , all
local regex_pattern "^(?:`regex_stubs')\("
gen numeric_calculate = ustrregexm(calculation, "`regex_pattern'")

local regex_stubs_2 : subinstr local numeric " " "|", all
local regex_pattern "^`regex_stubs_2'$"
gen numeric_force = ustrregexm(name, "`regex_pattern'")

local regex_stubs_3 : subinstr local string " " "|", all
local regex_pattern "^`regex_stubs_3'$"
gen string_force = ustrregexm(name, "`regex_pattern'")

label define question_type_M 1 "String" 2 "Select One" 3 "Select Multiple" ///
	4 "Numeric" 5 "Date" 6 "Datetime" 7 "GPS" ///
	-111 "Group Boundary" -222 "Note" ///
	-333 "Geopoint" -555 "Other" 
	
gen question_type=.

label values question_type question_type_M

replace question_type = 1 if inlist(type, "text", "deviceid", "image") ///
	| preloaded==1 ///
	| (type == "calculate" & numeric_calculate == 0) ///
	| string_force == 1
	
replace question_type = 2 if word(type, 1) == "select_one" ///
	& missing(question_type)

replace question_type = 3 if word(type, 1) == "select_multiple"

replace question_type = 4 if (!inlist(type, "date", "text") & missing(question_type)) ///
	| numeric_force == 1

replace question_type = 5 if inlist(type, "date", "today")

replace question_type = 6 if inlist(type, "start", "end", "submissiondate")

replace question_type= 7 if inlist(type, "geopoint", "geotrace") 

replace question_type = -111 if inlist(type, "begin_group", "end_group", ///
	"begin_repeat", "end_repeat")
replace question_type = -222 if note==1
replace question_type = -333 if type == "text audit"
replace question_type = -555 if missing(question_type)

drop if question_type < -111

// Create a new variable called "order" to retain the original order of variables in the instrument
gen order = _n
	
*===============================================================================
* 	Make Groups Dataset
*===============================================================================

local n_groups = 0
local n_repeats = 0
local grouplist 0
local r_grouplist 0
local group_label 0 "Survey-level"
local repeat_label 0 "Survey-level"

gen group = 0
gen repeat_group = 0

sort order
forvalues i = 1/`c(N)' {
	
	local j = `i' + 1
	
	if type[`i'] == "begin_group" { 
		
		local ++n_groups
		local grouplist = strtrim(stritrim("`grouplist' `n_groups'"))
		local group_label `group_label' `n_groups' "`=labelStata[`i']'"
		
		local type = question_type[`j']
		while `type' == 7 {
			
			local ++j
			local type = question_type[`j']
			
		}
		
	}
		
	else if type[`i'] == "begin_repeat" {
		
		local ++n_repeats
		local r_grouplist = strtrim(stritrim("`r_grouplist' `n_repeats'"))
		local repeat_label `repeat_label' `n_repeats' "`=labelStata[`i']'"
		
		local type = question_type[`j']
		while `type' == 7 {
			
			local ++j
			local type = question_type[`j']
			
		}
		
	}
	
	else if type[`i'] == "end_group" {
		
		local grouplist = strtrim(stritrim("`grouplist'"))
		local grouplist = ///
			substr("`grouplist'", 1, length("`grouplist'") ///
			- strlen(word("`grouplist'", -1)))
		
	}
	
	else if type[`i'] == "end_repeat" {
		
		local r_grouplist = strtrim(stritrim("`r_grouplist'"))
		local r_grouplist = ///
			substr("`r_grouplist'", 1, length("`r_grouplist'") ///
			- length(word("`r_grouplist'", -1)))
		
	}
	
	local current_group = word("`grouplist'", -1)
	local current_repeat = word("`r_grouplist'", -1)
	
	replace group = `current_group' in `i'
	replace repeat_group = `current_repeat' in `i'
	
}

clonevar tex_varlabel = labelEnglishen
clonevar tex_hint = hint
foreach var of varlist tex_varlabel tex_hint {
	
	replace `var' = ustrto(`var', "ascii", 1)
	foreach i in 92 35 36 37 38 95 94 123 125 126 {

		replace `var' = subinstr(`var', `"`=char(`i')'"', `"\\`=char(`i')'"', .)
		
	}
	
}

*===============================================================================
* 	Choices
*===============================================================================

// Create a new frame called "choices" to hold the value labels
tempname choices
frame create `choices' 
cwf `choices' 

// Import the "choices" sheet from the instrument Excel file specified by the "instname" variable
import excel "`instname'", firstrow clear sheet(choices)

// Rename the "listname" variable to "list_name" for consistency
cap rename listname list_name 

// Keep only the "list_name", "name", and "label" variables
keep list_name name label 

// Remove any empty rows from the dataset
missings dropobs, force 
tostring name, replace

foreach var of varlist list_name name label {
	
	replace `var' = strtrim(stritrim(`var'))
	
}

// Remove any rows where the "name" variable is not a number (i.e. non-standard labeling)
drop if !regexm(name, "^[0-9\-]+$") 

// Create a new variable called "order" to retain the original order of variables in the instrument
gen order = _n 

// Create a clone of the "name" variable for programming purposes
clonevar name1 = name 

// Replace any minus signs in the "name" variable with underscores, to match how SurveyCTO handles value labels
replace name = subinstr(name, "-", "_", 1) 

// Replace any dollar signs in the "label" variable with pound signs, to prevent Stata from parsing them as macros
replace label = subinstr(label, "$", "#", .)

// Remove any double quotes from the "label" variable
replace label = subinstr(label, `"""', "", .)
replace label = subinstr(label, char(10), "", .)

// Remove any spaces from the "list_name" variable
replace list_name = subinstr(list_name, " ", "", .)

// Create a local macro called "brek" containing a line break character
local brek = char(10) 
local tab = char(9)

// Remove any line breaks from the "name" and "name1" variables
foreach var of varlist name name1 {
	replace `var' = subinstr(`var', "`brek'", "", .)
}

*===============================================================================
* 	Variables
*===============================================================================

cwf `qs'
compress
drop if question_type < 0 | ///
	preloaded == 1 | ///
	inlist(type, "calculate", "today", "start", "end", "deviceid")
sort order

label define repeat_label `repeat_label'
label values repeat_group repeat_label

label define group_label `group_label'
label values group group_label

gen first_repeat = repeat_group[_n] != repeat_group[_n - 1] | ///
	missing(repeat_group[_n - 1])
gen first_group = group[_n] != group[_n - 1] | ///
	missing(group[_n - 1])

gen within = repeat_group[_n + 1] if repeat_group[_n + 1] < repeat_group[_n]
bysort repeat_group: ereplace within = max(within)
sort order

// generate some helpful macros
local brek = char(10)
local tab = char(9)
foreach thing in brek tab {
	
	forvalues z = 2/3 {
	
		local `thing'`z' = `z' * "``thing''"
	
	}
	
}

local hbanner = "*" + ("=")* 65
local lbanner = "*" + ("-") * 65

cap file close mytex
file open mytex using "`texdirectory'/manuscript.tex", write replace 
write_preamble, title("`surveyname'")
forvalues i = 1/`c(N)' {
	
	local qtype = question_type[`i']
	local qtext = tex_varlabel[`i']
	local label = type2[`i']
	local hint = tex_hint[`i']
	
	if first_repeat[`i'] == 1 {
		
		local grlabel : label repeat_label `=repeat_group[`i']'
		file write mytex _n(2) `"\section{\bf{Repeat Group}: `grlabel'}"'
		
	}
	
	if first_group[`i'] == 1 {
		
		local grlabel : label group_label `=group[`i']'
		file write mytex _n(2) `"\subsection{`grlabel'}"'
		
	}
	
	file write mytex _n(2) `"\begin{SurveyQuestion}[`qtext']"'
	
	if `qtype' == 1 {
		
		file write mytex `"\textit{[Open \textbf{text} entry]}"'
		
	}
	else if inlist(`qtype', 2, 3) {
			
		frame `choices' {
			
			valuesof label if list_name == `"`label'"'
			if `"`r(values)'"' == "" {
				
				file write mytex _n `"\textit{[select from answer options]}"'
				
			}
			else {
				
				file write mytex _n `"\begin{enumerate}[label=\alph*)]"'
				foreach t in `r(values)' {
					
					file write mytex _n `"\item `t'"'
					
				}
				file write mytex _n `"\end{enumerate}"'
			}
			
		}

	}
	else if `qtype' == 4 {
		
		file write mytex `"\textit{[Open \textbf{number} entry]}"'
		
	}
	else if `qtype' == 5 {
		
		file write mytex `"\textit{[Select a \textbf{date}]}"'
		
	}
	else if `qtype' == 6 {
		
		file write mytex `"\textit{[Select a \textbf{date and time}]}"'
		
	}
	else if `qtype' == 7 {
		
		file write mytex `"\textit{[Record your \textbf{GPS position}]}"'
		
	}
	
	if `"`hint'"' != "" {
		
		file write mytex _n `"\\ \hint{`hint'}"'
		
	}
	
	file write mytex _n `"\end{SurveyQuestion}"'
	
}

file write mytex _n(2) `"\end{document}"'
file close mytex 


}
end

cap prog drop write_preamble
prog define write_preamble 
syntax, TITLE(string)
qui {
	
local hbanner = "*" + ("=")* 65
local lbanner = "*" + ("-") * 65
local brek = char(10)
local tab = char(9)
local skip = "///" + char(10) + char(9)
foreach thing in brek tab {
	
	forvalues z = 2/3 {
	
		local `thing'`z' = `z' * "``thing''"
	
	}
	
}
	
local date : display %tdDayname,_dd_Month_CCYY date("`c(current_date)'", "DMY")
local date `date'

file write mytex `"%%%%%%%%%%%%%%%%%%%%%%%%%%"' _n ///
	`"%%%"' _n ///
	`"%%%      Preamble"' _n ///
	`"%%%"' _n ///
	`"%%%%%%%%%%%%%%%%%%%%%%%%%%"' _n(2) ///
	`"\documentclass[11pt]{article}"' _n(2) ///
	`"%%% packages"' _n ///
	`"\usepackage[colorlinks=true, allcolors=blue, obeyspaces]{hyperref}"' _n ///
	`"\usepackage{amssymb, amsfonts, amsmath}"' _n ///
	`"\usepackage{bm}"' _n ///
	`"\usepackage[margin=1in]{geometry} % full-width"' _n ///
    `"\topskip        =   20pt"' _n ///
    `"\parskip        =   10pt"' _n ///
    `"\parindent      =   0 pt"' _n ///
    `"\baselineskip   =   15pt"' _n(2) ///
	`"\usepackage{setspace}"' _n ///
    `"\onehalfspacing"' _n ///
	`"\usepackage{adjustbox}"' _n ///
	`"\usepackage{rotating}"' _n ///
    `"\numberwithin{table}{section}"' _n ///
	`"\usepackage{longtable}"' _n ///
	`"\usepackage{booktabs}"' _n ///
	`"\usepackage{dcolumn}"' _n ///
    `"\newcolumntype{d}[1]{D{.}{.}{#1}}"' _n ///
	`"\usepackage{natbib}"' _n ///
    `"\bibliographystyle{plainnat}"' _n ///
	`"\usepackage{pdflscape}"' _n ///
	`"\usepackage{appendix}"' _n ///
	`"\usepackage{graphicx}"' _n ///
	`"\usepackage{float}"' _n ///
	`"\usepackage{tablefootnote}"' _n ///
	`"\usepackage{threeparttable}"' _n ///
	`"\usepackage{comment}"' _n ///
	`"\usepackage{csquotes}"' _n ///
	`"\usepackage{tikz}"' _n ///
	`"\usetikzlibrary{calc}"' _n ///
	`"\usepackage{caption}"' _n ///
	`"\captionsetup{width=.9\linewidth, font=small}"' _n ///
	`"\usepackage{etoolbox}"' _n(2) ///
	`"\newcommand{\sym}[1]{\rlap{#1}} % for the stars"' _n(2) ///
	`"\usepackage{enumitem} % Better list environment control"' _n ///
	`"\usepackage{alphalph} % For alphanumeric labeling"' _n(2) ///
	`"% Define a new list to handle extensive multiple choices"' _n ///
	`"\SetLabelAlign{CenterWithParen}{\makebox[1em]{#1)}}"' _n ///
	`"\newlist{choices}{enumerate}{1}"' _n ///
	`"\setlist[choices,1]{label=\AlphAlph{\value*},align=CenterWithParen}"' _n(2) ///
	`"% Custom environment for survey questions"' _n ///
	`"\newcounter{question}[subsection]"' _n ///
	`"\newenvironment{SurveyQuestion}[1][]{%"' _n ///
   `"\refstepcounter{question}%"' _n ///
    `"\par\medskip"' _n ///
    `"\noindent\textbf{\thequestion. #1} \rmfamily%"' _n ///
	`"\begin{choices}%"' _n ///
	`"}{%"' _n ///
    `"\end{choices}"' _n ///
    `"\medskip"' _n ///
	`"}"' _n(2) ///
	`"% Command to include hints"' _n ///
	`"\newcommand{\hint}[1]{%"' _n ///
    `"\par"' _n ///
    `"\small"' _n ///
	`"\noindent\textit{Hint: #1}"' _n ///
    `"\medskip"' _n ///
	`"}"' _n(2) ///
	`"%%%%%%%%%%%%%%%%%%%%%%%%%%"' _n ///
	`"%%%"' _n ///
	`"%%%     Front page"' _n ///
	`"%%%"' _n ///
	`"%%%%%%%%%%%%%%%%%%%%%%%%%%"' _n(2) ///
	`"\title{`title' Questionnaire}"' _n ///
	`"\author{`c(username)'}"' _n ///
	`"\date{`date'}"' _n(2) ///
	`"\begin{document}"' _n ///
	`"\maketitle"' _n(2) ///
	`"\begin{abstract}"' _n ///
	`"\centering"' _n ///
	`"This is the questionnaire produced for the `title' Survey."' _n ///
	`"\end{abstract}"' _n(2) ///
	`"\newpage"' _n(2) ///
	`"%%%%%%%%%%%%%%%%%%%%%%%%%%"' _n ///
	`"%%%"' _n ///
	`"%%%   Introduction"' _n ///
	`"%%%"' _n ///
	`"%%%%%%%%%%%%%%%%%%%%%%%%%%"' _n(2) ///
	`"\tableofcontents"' _n(2) ///
	`"\clearpage"' _n(2)
	
}

end
