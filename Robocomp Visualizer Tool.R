
############################### ROBOCOMP VISUALIZER TOOL #################################

# This tool generates graphs for each data stream and export it into excel file
# It automatically extracts all the case log and insert log files from their respective folders.
# For each case_id present in case_log, a excel file will be generated.
# Each excel file will have a graph, metadata and data of respective data stream. 

###########################################################################################

rm(list=ls())
gc()
options(stringsAsFactors = FALSE)

############ Directory Locations ##########################################################

setwd("K:\\Kamal\\R-Project\\Robocomp Visualizer Tool Ver1.1")
functions <- paste(getwd(), "/Functions/functions.R", sep = "")
wd_case_log<-"K:\\UDI\\2014\\RoboComp\\case_logs\\"                  # used in load_case_log(). Load case log directory
wd_insert_log<-"K:\\UDI\\2014\\RoboComp\\insert_logs_complete\\"     # used in load_insert_logs(). Load insert log directory
directory_name <- paste("K:/UDI/2014/RoboComp/visuals/",sep="")      # used in plot_graph(). Visual location

########### Source all functions from functions.R #########################################

source(functions)

########### Install packages if not installed. Load packages if installed #################

checkInstallRequireLibrary(package="ggplot2")
checkInstallRequireLibrary(package="plyr")
checkInstallRequireLibrary(package="excel.link")
checkInstallRequireLibrary(package="grid")
checkInstallRequireLibrary(package="data.table")

################# Import dq_metadata.csv file ###################################

# Extracting pulse multiplier value from dq_metadata.csv
metadata_filename <- paste(Sys.getenv("USERPROFILE"),"\\Desktop\\R\\metadata_utils\\dq_metadata.csv.gz",sep="")
metadata <- read.csv(file = metadata_filename)
metadata <- metadata[metadata$DS_TYPE %in% "METER", c("dsid","PULSE_MULTIPLIER")]

################# Load all case_logs and insert logs ############################

case_log <- load_case_logs()            # load all case_log files present in case_log folder
insert_log <- load_insert_logs()        # load all insert_log files present in insert_log_complete folder

################# Loop for each case_log ########################################

for(j in 1:nrow(case_log))
{
  if(case_log$Rdata_file[j]!="")
  {
    ifelse(file.exists(case_log$Rdata_file[j]) == TRUE, load(as.character(case_log$Rdata_file[j])), next) # Load Rdata file if it is available in case_log
  }
  else
    next         
  
  ifelse(case_log$mapping_type[j]!="", case_mapping_type <- as.character(case_log$mapping_type[j]), print("No Mapping found"))       # Describe Mapping type from case_log
  
  ################# Generate platform_list ###################################
  
  ifelse(length(insert_log$dsid[insert_log$Case_Id == case_log$Case_Id[j]])!= 0, platform_list<-insert_log$dsid[insert_log$Case_Id == case_log$Case_Id[j]], next)     # If insert log is present then Generate Platform List else skip this case log's case_id
  
  platform_list <- unique(c("acct_ds",platform_list))    # Combine acct_ds to platform_list
  
  ######################################################################################
  
  plot_graph(Rdata_file, platform_list, insert_log, case_mapping_type)    # call plot_graph function
}
