## Test CCAL lab unpivot.
## Version 1.0.

## Created by: Nick Bywater
## Government agency: National Park Service, CAKN
## License: The Unlicense (Public Domain). See: https://unlicense.org/.

## Purpose: This script tests the validity of the unpivot
## transformation performed on pivoted source CCAL lab data as
## performed by the Streams Water Chemistry Utility (a MS Access
## utility written in VBA). It re-pivots the unpivoted output produced
## by the VBA utility and then compares each row-ID/column-name value
## of this output against each row-ID/column-name value of the source
## CCAL lab data. It keeps a count of all performed value comparisons
## and returns these counts in a list. If any value comparisons fail,
## it prints out those problem comparisons.

library(data.table)
library(RODBC)

## This function compares the values of output table data values to
## the source CCAL table data values.
##
## 1. The pivot function 'FUN' pivots the upivoted output data.
## 2. It then matches the pivoted output column names to the source
##    data column names.
## 3. After matching the names, it removes source names that did not
##    match the pivoted output column names.
## 4. For each row of the pivoted output, it iterates over each output
##    column name, and for each of these column names it iterates over
##    the source column names until if finds a matching output column
##    name in the prefix of one of the source names.
## 5. If there is a match, the code attempts to compare the values of
##    the pivoted output and source data where they match on row ID
##    and column name.
compare_values <- function(source_pivot_dt, output_unpivot_dt, FUN=pivot_function) {
   output_pivoted_dt <- FUN(output_unpivot_dt)
   source_pivot_col_pos <- pmatch(colnames(output_pivoted_dt), colnames(source_pivot_dt))
   source_pivot_col_pos <- source_pivot_col_pos[!is.na(source_pivot_col_pos)]
   source_pivot_col_names <- colnames(source_pivot_dt)[source_pivot_col_pos]

   total_compare_count <- 0
   fail_compare_count <- 0
   na_compare_count <- 0
   string_compare_count <- 0
   num_compare_count <-0

   for (r_id in output_pivoted_dt$rownumber) {
       for (i in colnames(output_pivoted_dt)) {
           for (j in source_pivot_col_names) {
               if (startsWith(j, i)) {
                   v1 <- as.vector(output_pivoted_dt[r_id])[i]
                   v2 <- as.vector(source_pivot_dt[r_id])[j]

                   ## First, if both source 'v2' and target 'v1' are
                   ## NA, then set 'is_equal' to TRUE. We can't compare
                   ## two NAs, so we set the flag 'is_equal' to true if
                   ## both 'v1' and 'v2' are NA.
                   is_equal <- FALSE
                   if (is.na(v1) && is.na(v2)) {
                       ## print(paste("row id:", r_id, "na1: ", v1, "na2: ", v2))
                       is_equal <- TRUE
                       na_compare_count <- na_compare_count + 1
                   }

                   ## Second, if 'is_equal' is false, then check to see
                   ## if 'v1' or 'v2' are strings, or can be converted
                   ## to strings and compared.
                   ##
                   ## 1. Try to convert 'v1' and 'v2' to a string.
                   ##    If successful and they compare equal, then set
                   ##    'is_equal' to TRUE.
                   ## 2. If this fails, then ether 'v1' or 'v2' (but
                   ##    not both) might be NA, or numeric that don't
                   ##    compare because of their representation.
                   if (!is_equal) {
                       if (as.character(v1) == as.character(v2)) {
                           ## print(paste("row id:", r_id, "string1: ", v1, "string2: ", v2))
                           is_equal <- TRUE
                           string_compare_count <- string_compare_count + 1
                       }
                   }

                   ## Third, if either 'v1' and 'v2' are still not
                   ## equal, then check to see if they are numeric. The
                   ## failure of 'v1' and 'v2' (together) to be
                   ## neither NAs nor strings may be due to the
                   ## numbers having a different string
                   ## representations, such as '49.10' and '49.1'.
                   if (!is_equal) {
                       if (is.numeric(as.numeric(v1)) && is.numeric(as.numeric(v2))) {
                           if (as.numeric(v1) == as.numeric(v2)) {
                               ## print(paste("row id:", r_id, "number1: ", v1, "number2: ", v2))
                               is_equal <- TRUE
                               num_compare_count <- num_compare_count + 1
                           }
                       }
                   }

                   ## Fourth, if the variable 'is_equal'
                   ## is still false, then none of the above
                   ## comarisons has succeeded.
                   ##
                   ## 1. Together, 'v1' and 'v2' are not both NA,
                   ##    numbers, or strings.
                   ## Or:
                   ## 2. It might be the case that 'v1' or 'v2' are
                   ##    different values.
                   ## Or:
                   ## 3. Some other comparison issue is at fault.
                   ##
                   ## Regardless, print off variables 'v1' and 'v2' to
                   ## indicate that there is a comparison issue.
                   if (!is_equal) {
                       print(paste("PROBLEM! row id: ",r_id,". [v1 = ", v1,"] [v2 = ", v2,"] Are equal?: ", is_equal, sep=""))
                       fail_compare_count <- fail_compare_count + 1
                   }

                   total_compare_count <- total_compare_count + 1
               }
           }
       }
   }
   ret <- list("total_compare_count"=total_compare_count,
               "fail_compare_count"=fail_compare_count,
               "na_compare_count"=na_compare_count,
               "string_compare_count"=string_compare_count,
               "num_compare_count"=num_compare_count,
               "total_value_count"=string_compare_count + num_compare_count,
               "all_values_equal"=if(fail_compare_count > 0) FALSE else TRUE)
   return(ret)
}

## Pivot the unpivoted ('stage_water_chemistry_...') water chemistry
## table (created by the VBA water chemistry utility) using the
## 'dcast' function of the data.table package. The pivot is defined on
## the row ID and the updated water chemistry names. So, a data table
## is created where the left column is row numbers and the pivoted
## columns are water chemmistry parameter names.
pivot_wc_table <- function(dt) {
    dt[, c("wc_parameter") := .(update_wc_parameter_names(dt))]
    dt <- dcast(dt, rownumber ~ wc_parameter, value.var="value")
    dt <- setnames(dt, names(dt), tolower(names(dt)))
    return(dt)
}

## Update the water parameter names with prefixes 'Duplicate' or
## 'Triplicate' if necessary.
update_wc_parameter_names <- function(dt) {
    updated_wc_param <- unlist(Map(update_wc_parameter_name,
                                   dt$analysisorder,
                                   dt$waterchemistryparameter), use.names=FALSE)
    return(updated_wc_param)
}

## This function uses the switch statement to return one of three
## water parameter name values.
## If acode = 1, return the wcparam without a prefix.
## If acode = 2, return the wcparam with 'Duplicate' prefix.
## If acode = 3, return the wcparam with 'Triplicate' prefix.
update_wc_parameter_name <- function(acode, wcparam) {
    return(switch(acode,
                  wcparam,
                  paste0("Duplicate ", wcparam),
                  paste0("Triplicate ", wcparam)))
}

## Read the unpivoted data produced by the VBA water chemistry
## utility.
read_unpivoted_data <- function(conn, table_name) {
    str_sql <- paste0("SELECT RowNumber, LabName, WaterChemistryParameter, Value, AnalysisDate, AnalysisOrder ",
                     "FROM [", table_name, "]")

    dt <- as.data.table(sqlQuery(conn, str_sql))
    dt <- setnames(dt, names(dt), tolower(names(dt)))
    return(dt)
}

## Read the pivoted source CCAL lab data, imported as a table in MS
## Access.
read_pivoted_data <- function(conn, table_name) {
    str_sql <- paste0("SELECT * ",
                     "FROM [", table_name, "]")

    dt <- as.data.table(sqlQuery(conn, str_sql))
    dt <- setnames(dt, names(dt), tolower(names(dt)))
    return(dt)
}

main <- function() {
    SOURCE_DB <- "c:/fake/directory/StreamsWaterChemistryUtility_v1.1.accdb"
    conn <- odbcConnectAccess2007(SOURCE_DB)

    unpivot_dt <- read_unpivoted_data(conn, "stage_water_chemistry_2024-12-16T13:51:51")
    pivot_dt <- read_pivoted_data(conn, "source_ccal_table")

    result <- compare_values(pivot_dt, unpivot_dt, FUN=pivot_wc_table)

    odbcClose(conn)

    return(result)
}
