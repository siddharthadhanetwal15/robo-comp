##functions.R
################################

## checkInstallRequireLibrary
## This function will check whether the required libraries are installed or not. 
## If package is not installed, it will install. Else package will be loaded.
checkInstallRequireLibrary <- function(packageName) 
{
  if(packageName %in% rownames(installed.packages()) == FALSE) {install.packages(packageName)}    
  library(packageName, character.only=TRUE)
}


## load_case_logs
## This function loads all case_ids of all case logs into a single data frame

load_case_logs<-function()
{
  #depends on fread and plyr packages  
  case_output_logs         <- list.files(path=toString(wd_case_log), full.names=FALSE, recursive=TRUE, include.dirs=FALSE)
  case_log_files   <- list.files(wd_case_log, pattern="*.csv$")
  
  #loop through case log files and aggregate to one data frame
  for (c in case_log_files) 
  {
    temp_case_log <- data.frame(data.table::fread(paste(wd_case_log, c, sep="\\")))
    temp_case_log$create_date <- as.POSIXct(substring(unlist(strsplit(c, "RoboComp_case_output_log_"))[2], 1, 13), format = "%Y%m%d_%H%M")
    if(!exists("case_log")) case_log <- temp_case_log else case_log <- rbind.fill(case_log, temp_case_log) 
  }
  case_log                 <- case_log[case_log$comment == "Insert File Generated",]
  #find the newest case log entry for each case
  case_log <- case_log[order(case_log$create_date, decreasing=T), ]
  #order case log on date and then deduplicate so newest is kept
  case_log <- case_log[!duplicated(case_log$Case_Id), ]
  rm(temp_case_log)
  
  return(case_log)
}

## load_insert_logs
## This function loads all case_ids of all insert logs into a single data frame
load_insert_logs<-function()
{
  insert_log_files <- list.files(wd_insert_log, pattern="*.csv$")
  
  #loop through insert files and aggregate to one data frame
  for(i in insert_log_files) {
    
    temp_insert_log <- data.frame(data.table::fread(paste(wd_insert_log, i, sep="\\")))
    
    temp_insert_log$log_date <- as.POSIXct(substring(unlist(strsplit(i, "RoboComp_insert_output_log_"))[2], 1, 13), format = "%Y%m%d_%H%M")
    
    if(!exists("insert_log")) insert_log <- temp_insert_log else insert_log <- rbind.fill(insert_log, temp_insert_log)
    rm(temp_insert_log)
    
  }
  insert_log <- insert_log[as.numeric(insert_log$log_date) %in% case_log$create_date, ]
  insert_log <- insert_log[order(insert_log$log_date, decreasing=T), ]
  return(insert_log)
}


#list all insert files (possibly add some filtering here later)

## plot_graph
## This function creates an account level view with an account comparison and individual meter level comparisons

plot_graph <- function(int_data, platform_list, insert_log, case_mapping_type) 
{
  
  case_id <- case_log$Case_Id[j]    
  operator_name <- case_log$Operator_formula__c[j]
  site_name <- case_log$Site_Name[j]
  c_date <- case_log$create_date[j]
  c_date <- format(c_date, format="%Y%m%d_%H%M")
  excel_file_name <- paste(directory_name,case_id,"_",operator_name,"_",site_name,"_","Robo Visual_",c_date,".xlsx",sep="")      # create name of output excel file
  
  if(file.exists(excel_file_name)!=TRUE)              # If visual doesn't exist then create excel file else tool will stop
  {
    xl.workbook.save(filename = excel_file_name)        # Create a Excel workbook
    
    sheets <- xl.sheets()                               # Get names of preloaded sheets of Workbook
    
    # loop through dsids to create a graph for the account level and each datastream
    
    for (dsid in platform_list) {
      
      if (dsid == "acct_ds"){
        
        util_meter <- "acct_util"
        util_subset <- util_meter
        
        
      } else {
        
        #define the utility meter that was used, this will be used as the label
        util_meter <- insert_log$Util_Id[insert_log$dsid==dsid&insert_log$Case_Id==case_id]
        
        #define what utility meter label will be used to subset the data
        #if 1:1 use the utility meter
        
        if(case_mapping_type=="1:1 Mapping"|case_mapping_type=="N:N Mapping"|
             (case_mapping_type=="N:M Mapping" & insert_log$Correlation[insert_log$dsid==dsid]!="Partial Sum")) util_subset <- insert_log$Util_Id[insert_log$dsid==dsid&insert_log$Case_Id==case_id]
 
        #if 1:n use ghost meter acct level, if n:1 and ratio  was done use parsed ghost meter level data
        if(case_mapping_type=="1:N Mapping"|
             case_mapping_type=="N:1 Mapping - Weighted Ratio"|
             case_mapping_type=="N:1 Mapping - Blanket Ratio"|
             (case_mapping_type=="N:1 Mapping - Subtraction" & insert_log$Correlation[insert_log$dsid==dsid]=="DS Requires Insert by Subtraction")) util_subset <- insert_log$Ghost_Util[insert_log$dsid==dsid]
        
        #if substraction was done the utility data for the good meter is equal to the platform data, if it is for the "bad" meter it will be equal to the ghost utility meter    
        if(case_mapping_type=="N:1 Mapping - Subtraction" & insert_log$Correlation[insert_log$dsid==dsid]=="No Insert Necessary - DS Used for Subtraction") util_subset <- dsid
        
        #if N:M mapping and not a 1:1 mapping use the ghost meter
        if(case_mapping_type=="N:M Mapping" & insert_log$Correlation[insert_log$dsid==dsid]=="Partial Sum") util_subset <- insert_log$Ghost_Util[insert_log$dsid==dsid]
        
      }
      #define platform and utility data and labels
      platform_data <- int_data[int_data$Id == dsid, ]
      if(util_subset!= "NO UTIL DATA") utility_data <- int_data[int_data$Id == util_subset, ] else utility_data <- NULL
      utility_data <- utility_data[!(rowSums(is.na(utility_data))==NCOL(utility_data)),] #delete NA rows. error occured at 5001300000kuFZSAA2
      platform_label <- dsid
      utility_label  <- util_meter
      
      png_file<-paste(dsid,".png",sep="")       # create a png file of graph
      png(filename=png_file, 
          type="cairo",
          units="in", 
          width=11.2, 
          height=8, 
          pointsize=25, 
          res=96)
      ## plot graph
      
      p <- create_plot(utility_data, platform_data, utility_label, platform_label)
      
      all_data <- create_plot(utility_data, platform_data, utility_label, platform_label)
      
      print(p)
      
      all_data <- create_dataframe(utility_data,platform_data)
      
      graph = current.graphics(filename=png_file)      # copies png file to excel file
      
      dev.off()
      
      xl.sheet.add(xl.sheet.name = dsid)        # Add sheet to workbook 
      # add metadata to excel sheet
      if(dsid=="acct_ds")
      {
        xlc[a1]="Util Id"
        xlc[a2]=util_meter
        xlc[b1]="DSID"
        xlc[b2]=dsid
        xlc[c1]="start_date_time "
        xlc[c2]=as.character(min(int_data$timestamp_local))
        xlc[d1]="end_date_time"
        xlc[d2]=as.character(max(int_data$timestamp_local))
      }
      else
      {
        xlc[a1]="Util Id"
        xlc[a2]=unique(util_meter)
        xlc[b1]="DSID"
        xlc[b2]= as.numeric(dsid)
        xlc[c1]="Pulse Multiplier"     
        xlc[c2]= unique(metadata$PULSE_MULTIPLIER[metadata$dsid==dsid])
        xlc[d1]="Correlation"
        xlc[d2]=unique(insert_log$Correlation[insert_log$dsid==dsid & insert_log$Case_Id==case_log$Case_Id[j]])
        xlc[e1]="Weighted Ratio"
         df_wt_ratio<-all_data[complete.cases(all_data[,c("p_demand","u_demand")]),]
        xlc[e2]=ifelse(is.na(sum(df_wt_ratio$p_demand*df_wt_ratio$u_demand)/sum(df_wt_ratio$u_demand)),"No Platform Data",sum(df_wt_ratio$p_demand*df_wt_ratio$u_demand)/sum(df_wt_ratio$u_demand))
        xlc[f1]="Display Label"
        xlc[f2]=unique(insert_log$display_label[insert_log$dsid==dsid& insert_log$Case_Id==case_log$Case_Id[j]])
        xlc[g1]="DS TZ"
        xlc[g2]=unique(insert_log$time_zone[insert_log$dsid==dsid & insert_log$Case_Id==case_log$Case_Id[j]])
        xlc[h1]="Aggregation"
        xlc[h2]=unique(insert_log$time_interval[insert_log$dsid==dsid & insert_log$Case_Id==case_log$Case_Id[j]])
        xlc[i1]="Num Intervals Inserted"
        xlc[i2]=unique(ifelse(insert_log$intervals_inserted[insert_log$dsid==dsid]=="", "NO INSERT NECESSARY",insert_log$intervals_inserted[insert_log$dsid==dsid & insert_log$Case_Id==case_log$Case_Id[j]]))
        xlc[j1]="Data Sparsity"
        xlc[j2]=unique(insert_log$data_sparsity[insert_log$dsid==dsid & insert_log$Case_Id==case_log$Case_Id[j]])
        xlc[k1]="start_date_time"
        xlc[k2]= as.character(min(int_data$timestamp_local))
        xlc[l1]="end_date_time"
        xlc[l2]=as.character(max(int_data$timestamp_local))
        xlc[m1]="Insert File"
        xlc[m2]=unique(insert_log$insert_file[insert_log$dsid==dsid & insert_log$Case_Id==case_log$Case_Id[j]])
      }  
      xl[a4]=graph                                    # Paste graph to Excel Sheet at a1 column 
      xlc[s2]="Data according to DSID"
      xlc[s3]=all_data[,c("timestamp","p_demand","p_estimated","p_insert","u_demand","u_flags")]                        # add actual data to excel sheet
      unlink(png_file)                                # delete png file
    }
    
    xl.sheet.delete(xl.sheet = sheets)                # delete preloaded sheets i.e. sheet1,sheet2,sheet3
    
    xl.workbook.save(filename = excel_file_name)      # save excel sheet with graphs
    
    xl.workbook.close()                               # close workbook
    
    updateVisualECRM(excel_file_name)
  
  }
  else
  {
    return
  }
}

##  create_plot
create_plot <- function(utility_data, platform_data, utility_label, platform_label){
  
  #define min and max of graph
  util_plat_bind <- rbind.fill(utility_data, platform_data)
  util_plat_bind <- util_plat_bind[!is.na(util_plat_bind$demand_kWh), ]
  y_min <- as.numeric(min(util_plat_bind$demand_kWh))
  y_min <- if(y_min < 0) y_min <- y_min*1.15 else y_min <- 0 #if less than zero make  
  y_max <- as.numeric(max(util_plat_bind$demand_kWh))*2
  
  #create a data frame of all relevant data to be plotted
  all_data <- create_dataframe(utility_data,platform_data)
  #define min and max of x
  x_min <- min(all_data$timestamp) - 4*60*60
  x_max <- max(all_data$timestamp) + 4*60*60
  
  #plot utility, platform, flags, and estimations
  #show_guide
  p <- ggplot(data=all_data, aes(x =  timestamp, y=blank))
  
  if(!(sum(is.na(all_data$u_demand))==nrow(all_data))) p <- p + geom_line(aes(y = u_demand,colour = "Utility Data"), size = .5, na.rm=T)#utility data
  
  if(!(sum(is.na(all_data$p_demand))==nrow(all_data))) p <- p + geom_line(aes(y = p_demand, colour = "Platform Data"), size = .5, na.rm=T) #platform data
  
  if (platform_label!="acct_ds" &!(sum(is.na(all_data$p_estimated_graph))==nrow(all_data))) p <- p + geom_line(aes(y = p_estimated_graph,colour = "Estimated"), size = .5, na.rm=T)
  
  if (platform_label!="acct_ds") p <- p + geom_vline(data=all_data,aes(xintercept= p_insert_graph),colour="grey52",alpha = 0.2, subset = .(p_insert_graph > 0),na.rm=T,show_guide=TRUE)
  
  if(!is.null(utility_data)) if (platform_label!="acct_ds") p <- p + geom_vline(aes(xintercept= u_flags_graph), color = "black", alpha = 0.2, subset = .(u_flags_graph > 0),show_guide=FALSE)
  
  if (platform_label!="acct_ds"&!(sum(is.na(all_data$p_estimated_graph))==nrow(all_data))) p <- p + geom_point(data = all_data, aes(y = p_estimated_graph,size="Estimated Point"), colour = "blue",na.rm=T)
  #scale by plot range
  y_plotrange <- c(y_min,y_max)
  x_plotrange <- c(x_min, x_max)
  p <- p + coord_cartesian(ylim = y_plotrange) + coord_cartesian(xlim = x_plotrange)
  
  #add labels and legends
  p <- p+scale_color_manual(values=c("Utility Data"="black","Platform Data"="red","Estimated"="Blue","Insert Flags"="grey52"))
  p <- p+theme(legend.text=element_text(size=9),legend.key.size=unit(0.8, "cm"),legend.title=element_blank(),legend.position="top",legend.direction = "horizontal")
  p <- p + xlab("Local Timestamp") + theme(axis.title.x = element_text(size = 9,face="bold",color="#3385FF")) + scale_x_datetime(breaks = "1 day") + 
    theme(text = element_text(size = 7))+theme(axis.text.x=element_text(angle=90,size=7,vjust=0.6))
  p <- p + ylab("Consumption (kWh)") + theme(axis.title.y = element_text(size = 9,face="bold",color="#3385FF"))
  if (utility_label=="acct_util" & platform_label=="acct_ds") p <- p + ggtitle("Account Level Comparison") else 
  p <- p + ggtitle(paste("DSID: ", platform_label, "     Utility Meter(s): ", utility_label, sep=""))
  p <- p + theme(plot.title=element_text(face="bold",size=10))
  
  #fix background
  p <- p + theme(panel.grid.minor=element_blank(), plot.title = element_text(size = 8), panel.background = element_blank())
  
  return(p)
 rm(utility_data)
 rm(platform_data)
  
}
##  create_dataframe
##  This function will construct a data frame according to dsid ### 
create_dataframe <- function(utility_data,platform_data)
{ 
  all_data <- data.frame("timestamp" = platform_data$timestamp_local, 
                         "p_demand" = as.numeric(platform_data$demand_kWh), 
                         "p_estimated_graph" = ifelse(platform_data$estimated==1, as.numeric(platform_data$demand_kWh), NA), 
                         "p_insert_graph" = ifelse(as.numeric(platform_data$insert_flag)==1, platform_data$timestamp_local, 0),                  
                         "p_estimated" = as.numeric(platform_data$estimated),
                         "p_insert" = as.numeric(platform_data$insert_flag))                 
  
  if(!is.null(utility_data)) ifelse(length(as.numeric(utility_data$demand_kWh[utility_data$timestamp_local == all_data$timestamp]))!=0,all_data$u_demand <- as.numeric(utility_data$demand_kWh[utility_data$timestamp_local == all_data$timestamp]),all_data$u_demand<-as.numeric(utility_data$demand_kWh)) 
  if(!is.null(utility_data)) ifelse(length(as.numeric(utility_data$util_flag[utility_data$timestamp_local == all_data$timestamp]))!=0 , all_data$u_flags  <- as.numeric(utility_data$util_flag[utility_data$timestamp_local == all_data$timestamp]),  all_data$u_flags<-as.numeric(utility_data$util_flag))
  if(!is.null(utility_data)) all_data$u_flags_graph  <- ifelse(as.numeric(all_data$u_flags)==1, all_data$timestamp, 0)
  return(all_data)
}

## This function will leave a comment about visual generated
updateVisualECRM <-function(excel_file_name)
{
  ## Update case page
  #     verbose <- F
  #     
  #     #### Source ####
  #     
  #     #### Sets R working directories ####
  #     wd_main <- Sys.getenv("wd_github_R")
  #     if (!file.exists(wd_main)) wd_main <- paste(Sys.getenv("USERPROFILE"),"\\Desktop\\R",sep="")
  #     if (!file.exists(wd_main)) stop("You need to define a system variable called 'wd_github_R' before running this script.
  #                                     \n This variable should be set to the directory containing R-related github repositories.
  #                                     \n To define it, go to Control Panel > System and Security > System > Advanced System Settings > Environment Variables > System Variables")
  #     username         <- Sys.getenv("USERNAME")
  #     wd_meterdu       <- paste(wd_main, "meter-data-utils", sep="\\")
  #     wd_metadu        <- paste(wd_main, "metadata_utils", sep="\\")
  #     
  #     #### Loads functions and extra packages ####
  #     sapply(list.files(wd_meterdu,pattern="*.R$", full.names=TRUE, ignore.case=TRUE), source, .GlobalEnv)
  #     library(zoo)
  #     
  #     
  #     
  #     ####Require the GUI elements ####
  #     needed.packs <- c("cairoDevice",
  #                       "gWidgets",
  #                       "gWidgetsRGtk2")
  #     
  #     sapply(needed.packs, function(x) if (!(x %in% installed.packages())) try(install.packages(x)))
  #     sapply(needed.packs, library,character.only=TRUE)
  #     
  #     options(guiToolkit="RGtk2")
  #     
  #     #### ECRM log in ####
  #     session          <- ecrm_login()
  #     while(class(session)[1]=="gButton"){}
  #   objectName <- "CaseComment"
  #   ParentId <- case_log$Case_Id[j]
  #   fields <- c(CommentBody=paste("Visual Generated \n File Location : ",excel_file_name))
  #   rforcecom.update(session, objectName, ParentId, fields)
}