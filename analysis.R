## -----------------------------------------------------------------------------
# install.packages('IRkernel')
# IRkernel::installspec()


## -----------------------------------------------------------------------------
# install.packages("remotes")
# install.packages("lintr")
# install.packages("vscDebugger")
# remove.packages("lintr")
# lintr::use_lintr(type = "tidyverse")
# install.packages("reshape2")
# install.packages("glue")
# install.packages("svglite")
# install.packages("rsvg")
# update.packages(ask = FALSE,c("ggplot2"))
# install.packages("vscDebugger")
#install.packages("ggpattern")
# install.packages(c("rmarkdown","knitr","xfun"))


## -----------------------------------------------------------------------------
# library(rmarkdown)
# library(knitr)
# library(xfun)
convert_ipynb_to_r <- function(
    input, 
    output = xfun::with_ext(input, "R"),
    keep_rmd = FALSE,
    ...
    ) {
        ## Check if necessary packages are installed
        if (!require("rmarkdown")) return("Missing necessary package: 'rmarkdown'")
        if (!require("knitr")) return("Missing necessary package: 'knitr'")

        ## Check if file extension is matches Jupyter notebook.
        if (tolower(xfun::file_ext(input)) != "ipynb") { 
            return( "Error: Invalid file format" )
        }

        ## Conversion process: 
        ## .ipynb to .Rmd
        rmarkdown::convert_ipynb(input)
        ## .Rmd to .R
        knitr::purl(xfun::with_ext(input, "Rmd"), output = output)

        ## Keep or remove intermediary .Rmd
        if (keep_rmd == FALSE) {
            file.remove(xfun::with_ext(input, "Rmd"))
    }
}
#input <- "analysis.ipynb"
#convert_ipynb_to_r(input)


## -----------------------------------------------------------------------------
#getwd()


## -----------------------------------------------------------------------------
# sessionInfo()
R.version.string


## -----------------------------------------------------------------------------
library(vscDebugger)


## -----------------------------------------------------------------------------
library(readxl)
library(ggplot2)
library(ggpattern)
library(gtsummary)
library(tidyverse)
library(gt)
library(glue)
library(svglite)
library(rsvg)


## -----------------------------------------------------------------------------
library(reshape2)


## -----------------------------------------------------------------------------
library(survival)
library(survminer)
library(dplyr) # %>%
library(plyr)
library(flextable)
library(tables)
library(crosstable)
library(IRdisplay)
library(rlang)
library(officer)


## -----------------------------------------------------------------------------
library(officer)
disp <- IRdisplay::display_html
set_flextable_defaults(
  font.family = "Open Sans",
  font.size = 10, 
  theme_fun = theme_vanilla,
  padding = 6,
  background.color = "#EFEFEF")
set_flextable_defaults(
  digits = 2,
  decimal.mark = ".",
  big.mark = ",",
  na_str = "." 
)  


## -----------------------------------------------------------------------------
as_flextable <- flextable::as_flextable


## -----------------------------------------------------------------------------
setwd('E:/SynologyDrive/Data/CarotidData2023')


## -----------------------------------------------------------------------------
shf <- read.csv('Analysis/recodeformulashortcuts.csv') # short functions
shf_1 <- shf[shf['type']==1,]
shf_2 <- shf[shf['type']==2,]


## -----------------------------------------------------------------------------
shf[shf['type']==1,]


## -----------------------------------------------------------------------------
report_file <- function(){return(format(Sys.time(), "%Y%m%d_%H%M%S"))}


## -----------------------------------------------------------------------------
list.files(path = getwd())


## -----------------------------------------------------------------------------
histogram_theme <- theme(
    plot.title = element_text(size = 8),# title
    axis.text = element_text(size = 4) ,# ticks
    axis.title = element_text(size = 6), # title of axis
    panel.background = element_rect(fill = "white"),
    axis.line = element_line(color = "black"),
  )


## -----------------------------------------------------------------------------
histogram_bin_label <- geom_text(
  stat = "count",
  aes(label =  after_stat(count)),
  vjust = -0.05,
  color = "black",
  size = 3
)


## -----------------------------------------------------------------------------
getwd()


## -----------------------------------------------------------------------------
# data_loaded <- read_csv( 'CAS_International_PROC_20230801.csv')
# data_ltf <- read_csv( 'CAS_International_LTF_20230801.csv')


## -----------------------------------------------------------------------------
# original<- read_excel('Analysis/recode.xlsx',sheet = "original") 
# recode <- read_excel('Analysis/recode.xlsx',sheet = "recode")
# analysis <- read_excel('Analysis/recode.xlsx',sheet = "analysis")
# recode <- dplyr::bind_rows(original,recode)
# recode <- filter(recode,!is.na(nvar))


## -----------------------------------------------------------------------------
# names(recode)[duplicated(names(recode))]


## -----------------------------------------------------------------------------

replace_shorts <- function(recode,fx,var_name,df,patterns,return_fx=FALSE,quote_var= FALSE) {
  # print(patterns)
  df_name<-deparse(substitute(df))
  # fx<-row[['fx']]
  # var_name<-row[['nvar']]
  
  # first replace #df with df value
  fx <- str_replace_all(fx,"\\$df",df_name)
  if(is.na(fx)){
    return(NA)
  }
  # parenthesis functions

  matches <- str_match_all(fx, patterns[1])
  m <- matches[[1]] # first element in match list 
  if (length(m) > 0) {
    for (i in 1:nrow(m)){
      mm<- str_match( m[i,],'^(.*?)([\\$\\#].*)')
      f <- shf_2[shf_2['var']==mm[[2]],'nvar']
      v <- mm[[3]]
      r <- paste(f,'(',v,')')  
      r<-gsub(" ","",r)
      r<-paste(' ',r)
      p <- str_c("\\b", mm[[2]], "\\",mm[[3]],"\\b")
      fx <- str_replace_all(fx,p,r)
    }
  } 
  
  # print(paste('1:',fx))
  # variables 
  matches <- str_match_all(fx, patterns[2])
  m <- matches[[1]] 
  if (length(m) > 0) {
    for (i in 1:nrow(m)){
      sign <- substr(m[i,],1,1)
      v <-substring( m[i,],2)
      is_only_digits <- grepl("^\\d+$", v)
      if(is_only_digits){
        r1<-paste(recode[recode['id']==v,'nvar'])
      } else {
        r1<-v
      }
      r2<- ifelse(quote_var == TRUE, paste('"',r1,'"'),r1)
      r3 <- ifelse(sign == "#", paste(df_name, "$", r2), r2)
      r4<-gsub(" ","",r3)
      r<-paste(' ',r4)
      p <- str_c("\\",sign,v,"\\b")
      fx <- str_replace_all(fx,p,r)      
    }
  } 

  # print(paste('2:',fx,'  v:',v,'  r1:',r1,'  r2:',r2,'  r3:',r3,'  r4:',r4,'  r:',r,'  p:',p))
  # non-parenthesis functions
  matches <- str_match_all(fx, patterns[3])
  m <- matches[[1]] 
  if (length(m) > 0) {
    for (i in 1:nrow(m)){
      v <- m[i,]
      r <- shf_1[shf_1['var']==v,'nvar']
      p <- str_c("\\b",v,"\\b")
      fx <- str_replace_all(fx,p,r)
    }
  } 
  # print(paste('3:',fx))
  if (return_fx ) {return(fx)}
  
tryCatch({
  tib <- tibble::tibble(!!var_name := eval(parse(text = fx)))
  df <- cbind(df,tib)
  print(paste('CREATED:',var_name,' with forumla ', fx))
}, error = function(e){
  print(paste('ERR:',var_name,' with forumla ', fx,' ERROR: ',e))
  stop()
}, finally = {
  return(list(df=df,fx=fx))
})
}



## -----------------------------------------------------------------------------
# arg<-"c($test,$14)"
# replace_shorts(recode=recode,fx=arg,var_name='',df=df,patterns=patterns,return_fx = TRUE,quote_var=TRUE) 


## -----------------------------------------------------------------------------
load_file <- function(file_name) {
  print(file_name)
  data_loaded <- read_csv(file_name)
  print(paste('loaded:',file_name, ',n=',nrow(data_loaded)))
  return(data_loaded)
}


## -----------------------------------------------------------------------------
reload_recode <- function(sub_pri,recode,original,patterns){
  # print(paste('reload_recode:sub_pri:',sub_pri,'recode:',recode,'original:',original,'patterns:',patterns))
  # print(paste('recode nrow',nrow(recode)))
  # print(recode)
  duplicates <- recode[duplicated(recode$nvar),]
  if (nrow(duplicates)>0){
    print(duplicates)
    stop(paste('ERROR :: DUPLICATES:',nrow(duplicates)))  
  }
  duplicates <- recode[duplicated(recode$id),]
  if (nrow(duplicates)>0){
    print(duplicates)
    stop(paste('ERROR :: DUPLICATES:',nrow(duplicates)))  
  }
  df<-sub_pri
  calc_columns <- recode[recode['f']=='fx',]

  for (i in 1:nrow(calc_columns)){
  # for (i in 1:1){
    result<-replace_shorts(recode=recode,fx=calc_columns[[i,'fx']],var_name=calc_columns[[i,'nvar']],df=df,patterns=patterns)
    df<-result$df
    fx<-result$fx
    recode$newfx[recode$var==calc_columns[[i,'nvar']]]<-fx
  }
  print(paste('reload_recode:n=',nrow(df)))
  return(list(df=df,recode=recode))
  }


## -----------------------------------------------------------------------------
ept <- function(t){
  return(eval(parse(text=t)))
}


## -----------------------------------------------------------------------------
get_args <-function(recode,tib_row,n_args,n__args,patterns,do_eval=TRUE,quote_var= TRUE){
# get_args
# row of tib 
# number of arguments to retreive
# parse first n arguments
# e.g get_args(analysis[[i]],3,2)
# return list of arguments
  args <- list()
  if (n_args > 0){
    for (i in 1:n_args){
      arg <- tib_row[[as.character(i)]]
      print(paste('arg:',arg))
        
      arg <- replace_shorts(recode=recode,fx=arg,var_name='',df=df,patterns=patterns,return_fx = TRUE,quote_var=quote_var)  
      arg <- gsub(" ","",arg)
      if (do_eval == TRUE) {  
        arg <- eval(parse(text = arg))
      }
      args[[i]]<- arg
    }
  }
  
  if (n__args > 0){
    for (i in 1:n__args){
        args[[n_args+i]]<- tib_row[[paste0("_",i)]]
    }
  }

  return(args) 
}


## -----------------------------------------------------------------------------
ctn1_shell <- function(recode,analysis_row,n_args,n__args,patterns,doc){
    processed_args<-get_args(recode,analysis_row,1,1,patterns)
    doc <- add_n(analysis_row[['in']],doc)
    ft<- ctn1(analysis_row[['in']],processed_args)
    ft<- labelit(ft, recode, processed_args[[1]])
    if(disp_){disp(to_html(ft))}
    doc <- add_text(doc, analysis_row[['title']])
    doc <- body_add_flextable(doc, ft)
    doc <- body_add_break(doc)
    return(doc)
}
ctn1 <- function(df_name,args){
  df<-get(df_name)
  cols<-args[[1]]
  arg<-args[[2]]

  ct <-  crosstable(
    df,
    cols =all_of(cols),
    percent_pattern =arg,
    percent_digits = 0
    
    )
    print(ct)
    ft<- ct |> 
    crosstable::as_flextable(spread_first_col = TRUE,separate_with = "label") |>
    # theme_tron() |> 
    fix_border_issues() |>
    autofit() |>
    set_caption(caption="Automated analysis")
    # print(ct)
  return(ft)
}


## -----------------------------------------------------------------------------
ctn2_shell <- function(recode,analysis_row,n_args,n__args,patterns,doc){
    processed_args<-get_args(recode,analysis_row,2,1,patterns)
    doc <- body_add_break(doc)
    doc <- add_n(analysis_row[['in']],doc)
    ft<- ctn2(analysis_row[['in']],processed_args)
    # ft<- labelit(ft, recode, processed_args[[2]])
    if (disp_) {disp(to_html(ft))}
    doc <- add_text(doc, analysis_row[['title']])
    doc <- body_add_flextable(doc, ft)
}

ctn2 <- function(df_name,args){
  df<-get(df_name)
  
  col<-args[[1]]
  cols<-args[[2]]
  arg<-args[[3]]
  ct <-  crosstable(
      data=df,
      cols =all_of(cols),
      by = all_of(col),
      percent_pattern =arg,
      percent_digits = 0,
      funs=c(mean,sd)
    )
    # # add frequencies to headers
    # frequencies <- table(df[[col]])
    # names(ct)<-lapply(names(ct),function(x){
    #   if(x %in% names(frequencies)){
    #     returning <- paste0(x,' (N=',frequencies[[x]],')')
    #     return (returning)
    #   }  else {
    #     return(x)
    #   }
    # })
    
    col_var <- recode[recode$nvar == col,'label'][[1]]

    
    print(tibble::tibble(ct))
    ft<- ct |> 
    crosstable::as_flextable(spread_first_col = TRUE,separate_with = "label") |>
    # theme_tron() |> 
    fix_border_issues() |>
    autofit() |>
    set_caption(caption=col) |>
    labelizor(part="header",
        labels=c("label"="variables","variable"="",setNames(col_var,col)),
        )
    # print(ct)
  return(ft)
}


## -----------------------------------------------------------------------------
ctn3_shell <- function(recode,analysis_row,n_args,n__args,patterns,doc){
        processed_args<-get_args(recode,analysis_row,1,1,patterns)
      doc <- add_n(analysis_row[['in']],doc)
      ft<- (ctn3(analysis_row[['in']],processed_args))$ft
      if(disp_){disp(to_html(ft))}
      doc <- add_text(doc, analysis_row[['title']])
      doc <- body_add_flextable(doc, ft)
      doc <- body_add_break(doc)
}
# CTN3  1 set of variables for crosstabling
ctn3 <- function(df_name,args){
  df<-get(df_name)
  
  vars <- args[[1]]
  stat <- args[[2]]
  var_levels <- list()
  var_level_labels <- list()
  var_count <- 0
  
  ctn <- function(df,col,cols,stat){
    ct <-  crosstable(
      df,
      cols =all_of(cols),
      by = all_of(col),
      percent_pattern =stat,
      percent_digits = 0
      ) 
      var_count <<- var_count +1 
      var_level_labels[[var_count]] <<- names(ct)
      
      l <<- length(names(ct))
      var_levels <<- c(var_levels, l-3)
      # print(vars(4:l))
    
      ct <- ct |>
        rename_at(vars(4:l), ~ paste0(col,"@", .))
      return(ct)
  }
  
  # first
  tib <- ctn(df,vars[1],vars[-1],stat) 
  if (length(vars) == 2) {
    return(tib)
  }
  
  # others
  for (i in seq_along(vars)[-1]) {
    vars_other <- vars[-i]
      ct <- ctn(df,vars[i],vars_other,stat)
      tib <- tib |> full_join (ct)
  }
  
  # replace NA with "."
  tib <- tib |> mutate_all(~ replace(.,  is.na(.), "."))
  
  # flextable start
  ft<-as_flextable(tib)
  ft<-ft|> delete_rows(i=1:2,part='header')
  
  # header row 2
  header_row_labels <- c("label","variable")
  col_widths <- c(1,1)
  for (i in seq_along(var_level_labels)) {
    var_level_label <- var_level_labels[[i]]
    n<- length(var_level_label)
    for (j in seq_along(var_level_label)[4:n]) {
      header_row_labels <- append(header_row_labels,var_level_label[j])
      col_widths <- append(col_widths,1)
    }
  }
  ft <- ft |> add_header_row(
      values = header_row_labels,
      colwidths = col_widths
    )
    
  # header row 1
  
  header_row_labels <- c("","")
  col_widths <- c(1,1)
  for (i in seq_along(vars)){
    header_row_labels <- append(header_row_labels,vars[i])
    col_widths <- append(col_widths,var_levels[[i]])
  }
  ft <- ft |> add_header_row(
      values = header_row_labels,
      colwidths = col_widths
    )
  ft <- ft |> 
  align(
    align = "center",
    part = "header"
  )  |>
  bold(
    bold = TRUE,
    part = "header"
  ) |>
  bold(
    j= 1:2,
    part = "body"
  )|> 
  align(
    align = "right",
    part = "body"
  )

  return(list(ft=ft,tib=tib))
}


## -----------------------------------------------------------------------------
ctn4_shell <- function(recode,analysis_row,n_args,n__args,patterns,doc){
        processed_args<-get_args(recode,analysis_row,2,1,patterns)
      doc <- add_n(analysis_row[[,'in']],doc)
      ft<- (ctn4(analysis_row[['in']],processed_args))$ft
      if(disp_){disp(to_html(ft))}
      doc <- add_text(doc, analysis_row[['title']])
      doc <- body_add_flextable(doc, ft)
      doc <- body_add_break(doc)
}
# CTN4 2 set of variables (cols and rows)
ctn4 <- function(df_name,args){
  df<-get(df_name)
  
  vars1 <- args[[1]] # columns
  vars2 <- args[[2]] # rows
  stat <- args[[3]]
  var_levels <- list()
  var_level_labels <- list()
  var_count <- 0
  ctn <- function(df,col,cols,stat){
    ct <-  crosstable(
      df,
      cols =all_of(cols),
      by = all_of(col),
      percent_pattern =stat,
      percent_digits = 0
      ) 
      var_count <<- var_count +1 
      var_level_labels[[var_count]] <<- names(ct)
      
      l <<- length(names(ct))
      var_levels <<- c(var_levels, l-3)
      ct <- ct |>
        rename_at(vars(4:l), ~ paste0(col,"@", .))
      return(ct)
  }
  
  # first
  tib <- ctn(df,vars1[1],vars2,stat) 
  if (length(vars1) == 1) {
    return(tib)
  }
  
  # others
  for (i in seq_along(vars1)[-1]) {
    ct <- ctn(df,vars1[i],vars2,stat)
    tib <- tib |> full_join (ct)
  }
  
  # replace NA with "."
  tib <- tib |> mutate_all(~ replace(.,  is.na(.), "."))
  
  # flextable start
  ft<-as_flextable(tib)
  
  # delete header rows
  ft<-ft|> delete_rows(i=1:2,part='header')
  
  # rebuild header row 2
  header_row_labels <- c("label","variable")
  col_widths <- c(1,1)
  for (i in seq_along(var_level_labels)) {
    var_level_label <- var_level_labels[[i]]
    n<- length(var_level_label)
    for (j in seq_along(var_level_label)[4:n]) {
      header_row_labels <- append(header_row_labels,var_level_label[j])
      col_widths <- append(col_widths,1)
    }
  }
  ft <- ft |> add_header_row(
      values = header_row_labels,
      colwidths = col_widths
    )
    
  # rebuild header row 1
  header_row_labels <- c("","")
  col_widths <- c(1,1)
  for (i in seq_along(vars1)){
    header_row_labels <- append(header_row_labels,vars1[i])
    col_widths <- append(col_widths,var_levels[[i]])
  }
  ft <- ft |> add_header_row(
      values = header_row_labels,
      colwidths = col_widths
    )
  ft <- ft |> 
  align(
    align = "center",
    part = "header"
  )  |>
  bold(
    bold = TRUE,
    part = "header"
  ) |>
  bold(
    j= 1:2,
    part = "body"
  )|> 
  align(
    align = "right",
    part = "body"
  )

  return(list(ft=ft,tib=tib))
}


## -----------------------------------------------------------------------------
refilter <- function(df_name,args){
  df<-get(df_name)
  arg <- args[[1]]
  df <- df |> filter(eval(parse(text = arg)))
  print(paste('filtered:',args,', n=',nrow(df)))
  return (df)
}


## -----------------------------------------------------------------------------
subset <- function(df_name,args){
  # print('SUBSET')
  df<-get(df_name)
  not_in_df <- args[!args %in% names(df)]
  if (length(not_in_df) > 0) {
    print(paste('subset:',not_in_df))
    stop(paste('NOT IN DF'))
  }
  # arg <- args[[1]]
  # print(args)
  # print(names(df[args]))
  df <- df[args]
  return (df)
}


## -----------------------------------------------------------------------------
summ_shell <- function(recode,analysis_row,n_args,n__args,patterns,doc){
        processed_args<-get_args(recode,analysis_row,2,1,patterns)
      doc <- body_add_break(doc)
      doc <- add_n(analysis_row[['in']],doc)
      ft<- summ(analysis_row[['in']],processed_args[[2]])
      if (disp_) {disp(to_html(ft))}
      doc <- add_text(doc, analysis_row[['title']])
      doc <- body_add_flextable(doc, ft)
}
summ <- function(df_name,args){
  df<-get(df_name)
  
  col<-args[[1]]
  cols<-args[[2]]
  filt = args[[3]]
  ft <- df |> 
  select(all_of(cols)) |>
  filter(eval(parse(text=filt)))|>
  summarizor(
    by =col,
    ) |> 
    as_flextable(
      sep_w =0, 
      spread_first_col = TRUE,
      columns_alignment ='right',
    ) |>
    theme_tron() |> 
    fix_border_issues() |>
    autofit() |>
    set_caption(caption=col) |>
    italic(j=1,part="body")
    # print(ft)
  return(ft)
}


## -----------------------------------------------------------------------------
freq1_shell <- function(recode,analysis_row,n_args,n__args,patterns,doc){
        processed_args<-get_args(recode,analysis_row,1,0,patterns)
      doc <- body_add_break(doc)
      doc <- add_n(analysis_row[['in']],doc)
      ft<- freq1(analysis_row[['in']],processed_args)
      if (disp_) {disp(to_html(ft))}
      doc <- add_text(doc, analysis_row[['title']])
      doc <- body_add_flextable(doc, ft)
}
freq1 <- function(df_name,args){
  df<-get(df_name)
  
  var <- args[[1]]
  ft <- proc_freq(df,var) |>
      theme_tron() |> 
    fix_border_issues() |>
    autofit() |>
    set_caption(caption=var)
  return(ft)
}


## -----------------------------------------------------------------------------
gt_glm_shell <- function(recode,analysis_row,n_args,n__args,patterns,doc){
  tib_row <- analysis_row
  args1<-get_args(recode,tib_row,1,0,patterns,do_eval=TRUE,quote_var=FALSE)
  arg1 <- regmatches(tib_row['1'], regexec("~(.*)", tib_row['1']))[[1]][2]
  temp<-regmatches(arg1, gregexpr("\\$\\w+", arg1))[[1]]
  # modified_temp <- sub("\\$", "", temp)
  tib_row['2']<-paste0("c(", paste(temp, collapse = ", "), ")")
  args2<-get_args(recode,tib_row,2,1,patterns,do_eval=FALSE)
  args3<-get_args(recode,analysis[i,],2,1,patterns,do_eval=FALSE)
  processed_args<-list(args1[[1]],args2[[2]],args2[[3]])

  doc <- body_add_break(doc)
  ft <- gt_glm(analysis_row[['in']],processed_args)
  if (disp_) {disp(to_html(ft))}
  doc <- add_text(doc, analysis_row[['title']])
  doc <- body_add_flextable(doc, ft)
  return(doc)
}
gt_glm <- function(df_name,args){
  df<-get(df_name)
  formula_obj <- formula(args[[1]])
  # print(formula_obj)
  model <- args[[3]]
  vars_list<-args[[2]]
  label_list <-lapply(ept(vars_list),make_match)
  # print(label_list)
  glm_obj <- glm(formula_obj,data=df,family=model)
  ft<- glm_obj |>
    tbl_regression(
      exponentiate = TRUE,
      label= label_list
    ) |>
    as_flex_table() 
  
  # Extract coefficients, standard errors, z-values, and p-values from the model summary
  # coefficients <- coef(summary(glm_obj))
  # se <- coefficients[, "Std. Error"]
  # z_values <- coefficients[, "z value"]
  # p_values <- coefficients[, "Pr(>|z|)"]
  # print(paste('coefficients',coefficients))

  # Create a data frame with coefficients, standard errors, z-values, and p-values
  # result_df <- data.frame(
  #   Coefficients = coefficients[, "Estimate"],
  #   `Std. Error` = se,
  #   `z value` = z_values,
  #   `Pr(>|z|)` = p_values
  # )
  # print(paste('results_df',result_df))
  # # Create a flextable from the data frame
  # ft <- flextable(result_df)
  
  return(ft)
}


## -----------------------------------------------------------------------------
freq2_shell <- function(recode,analysis_row,n_args,n__args,patterns,doc){
        processed_args<-get_args(recode,analysis_row,2,0,patterns)
      doc <- body_add_break(doc)
      doc <- add_n(analysis_row[['in']],doc)
      ft<- freq2(analysis_row[['in']],processed_args)
      if (disp_) {disp(to_html(ft))}
      doc <- add_text(doc, analysis_row[['title']])
      doc <- body_add_flextable(doc, ft)
}
freq2 <- function(df_name,args){
  df<-get(df_name)
  
  col <- args[[1]]
  row <- args[[2]]
  ft <- proc_freq(df,row,col) |>
      theme_tron() |> 
    fix_border_issues() |>
    autofit() |>
    set_caption(caption=col)
    print(ft)
  return(ft)
}


## -----------------------------------------------------------------------------
add_n <- function(df_name,doc,t=""){
  df<-get(df_name)
  doc <- doc %>%
    body_add_par(paste(t), pos = "after",style="Normal") |>
    body_add_par(paste("n=",nrow(df)),style="Normal", pos = "after") |>
    body_add_par("", pos = "after",style="Normal")
  return(doc)
}


## -----------------------------------------------------------------------------
add_text <- function(doc,t,level=2){
  if (level==1){
    text_style <- fp_text(font.size = 20,bold=TRUE)    
    par_style <- fp_par(text.align = "left")
  
  } else if (level==2){
    text_style <- fp_text(font.size = 14,bold=TRUE)   
    par_style <- fp_par(text.align = "left")
  
  } else if  (level>8){ # paragraph
    text_style <- fp_text(font.size = 12,bold=FALSE)
    par_style <- fp_par(text.align = "left")
  }
  doc <- body_add_fpar(doc, fpar( ftext(t, prop = text_style), fp_p = par_style ) )
  return(doc)
}


## -----------------------------------------------------------------------------
factorize_variables <- function(df,recode){
  # df<-get(df_name)
  values<-recode[!is.na(recode$values), c('nvar','values')]
  for (i in 1:nrow(values)){
    var <- values[[i,'nvar']]
    values_  <- values[[i,'values']]
    split_values  <- strsplit(values_,',')
    numeric_values <- as.list(sapply(split_values, function(x) as.numeric(sub("(.*)=.*", "\\1", x))))
    character_values <- as.list( sapply(split_values, function(x) sub(".*=(.*)", "\\1", x)))
    df[[var]]<-factor(df[[var]],levels=numeric_values,labels=character_values)
  }
  return(df)
}


## -----------------------------------------------------------------------------
relevel_variables <- function(df,recode){
  values<-recode[!is.na(recode$relevel), c('nvar','relevel','values')]
  for (i in 1:nrow(values)){
    var <- values[[i,'nvar']]
    relevel <- values[[i,'relevel']]
    values_  <- values[[i,'values']]
    split_values  <- strsplit(values_,',')
    numeric_values <- as.list(sapply(split_values, function(x) as.numeric(sub("(.*)=.*", "\\1", x))))
    character_values <- as.list( sapply(split_values, function(x) sub(".*=(.*)", "\\1", x)))
    if(length(which(numeric_values==relevel))>0){
      df[[var]]<-relevel(df[[var]],ref=character_values[[which(numeric_values == relevel)]])
    }
  }
  return(df)
}


## -----------------------------------------------------------------------------
labelit <- function(ft,recode,args){
  print(paste('labelit:args>',args))
  labels = (recode[recode$nvar %in% args, "label"])$label
  names(labels) = as.list(args)
  ft <- labelizor(
    x = ft, j = c("label"),
    labels = labels
  )
  return(ft)
}


## -----------------------------------------------------------------------------
rename_variables <- function(data_loaded,recode,original,args){
  recode <- dplyr::bind_rows(original,recode)
  recode <- filter(recode,!is.na(nvar))
  df<- data_loaded
  df_name <- args[[1]]
  # Rename variables using recode.csv
  recode_data <- function(r){
    names(df)[names(df) == r[['var']]] <<- r[['nvar']]
  }

  temp_recode <- recode |> filter(f==df_name)
  # print(nrow(temp_recode))
  for (i in 1:nrow(temp_recode)){
  # for (i in 1:1){
    r <- temp_recode[i,]
    recode_data(r)
  }
  sub_pri<-df[temp_recode[['nvar']]]

  return(list(sub_pri=sub_pri,recode=recode))
}


## -----------------------------------------------------------------------------
aggregate_count <- function(df_name,args){
  df<-get(df_name)
  var <- gsub(" ", "", processed_args[[1]])
  tib <- as_tibble(table(df[[var]]),.name_repair = "minimal")
  names(tib)[1]<-var
  return(tib)
}


## -----------------------------------------------------------------------------
ggplot_hist_l <- function(df_name,args,doc,titles){
  variable_names <- ept(args[[1]])
  print(variable_names)
  titles <-  strsplit(titles, split = "|", fixed = TRUE)[[1]]
  for (i in seq_along(variable_names)) {
    var <- variable_names[[i]]  # Get the variable name
    title <- titles[i]          # Get the corresponding title
    var_name<-deparse(substitute(var))
    print(var_name)
    plot<-ggplot_hist(df_name,c(var_name))
    doc <- doc |> body_add_break() |> add_text( title) |>body_add_gg(plot, width = 6, height = 4)
  }
  return(doc)
}


## -----------------------------------------------------------------------------
ggplot_hist <- function(df_name,args){
  # df<-get(args[[2]])
  df<-get(df_name)
  x_var <-ept(args[[1]])
  print(x_var)
  print(  names(df))
  histogram <- ggplot(df, aes(x = .data[[x_var]])) +
  geom_histogram(binwidth = 0.5, fill = "grey", color = "black", alpha = 0.5) +  # Customize the histogram aesthetics
  labs(x = "Values", y = "Frequency", title = "Histogram of Values") +  # Add labels and title
  stat_bin(geom = "text", aes(label = ifelse(..count.. > 0, ..count.., "")), vjust = -0.5) +  # Add count labels on top of each bar
  theme_minimal()  # Apply a theme (optional)
  return(histogram)
}


## -----------------------------------------------------------------------------
recode_list<-function(recode,analysis,i,doc){
  result<-lapply(analysis[['1']],function(x){
  matches <- regmatches(x,gregexpr("[\\$#]([0-9]+)",x))[[1]]
  if(length(matches) > 0) {
      return(matches)
    } else {
      return(NULL)
    }
  })
  found <- gsub("\\$","",result[sapply(result, function(x) !is.null(x))]|> 
    unlist()) |> 
    tibble()
  names(found)[1]<- 'id'
  temp_recode <- mutate(recode,id=as.character(id))
  working_recode<-left_join(found,temp_recode,by='id')

  doc <- add_text(doc, 'Selected Variables')
  # doc <- add_n(recode,doc,"variables count:")
  doc <- body_add_table(doc,working_recode[!is.na(working_recode$newfx),c("var","newfx")])
  doc <- body_add_break(doc)
  
  
   doc <- add_text(doc, 'All Variables')
  # doc <- add_n(recode,doc,"variables count:")
  doc <- body_add_table(doc,recode[!is.na(recode$newfx),c("var","newfx")])
  doc <- body_add_break(doc)
  return(doc)
}


## -----------------------------------------------------------------------------
set_widths <- function(ft,col1=2,other_cols=0.75){
  ft<-width(ft,j=1,width=col1)
  for(i in 2:ncol_keys(ft)){
    ft<-width(ft,j=i,width=other_cols)
  }
  return(ft)
}


## -----------------------------------------------------------------------------
gt_tbl_summary_shell <- function(recode,analysis_row,n_args,n__args,patterns,doc){
    processed_args<-get_args(recode,analysis_row,2,3,patterns)
    doc <- add_n(analysis_row[['in']],doc)
    font_size <- analysis_row[['font']]
    if (is.na(font_size)) {
      font_size <- 10
    }
    ft<- gt_tbl_summary(analysis_row[['in']],processed_args) |>
      set_widths(col1=2,other_cols=0.75) |>
      labelit( recode, processed_args[[1]]) |>
      fontsize(size=font_size)
    if (!is.na(analysis_row[['title']])) {  
      doc <- add_text(doc, analysis_row[['title']])
    }
    doc <- body_add_flextable(doc, ft,split=TRUE)
    doc <- body_add_break(doc)
    return(doc)
}
make_match <- function(x) {
  found<-(recode[recode$nvar == x, "label"])$label
  ept(paste0(x,"~ '",found, "'"))
}
match_single <- function(x) {
  (recode[recode$nvar == x, "label"])$label
}
keep_with_next <- function(ft) {
  indents <- unique(ft$body[[8]]$pars$padding.left$data[,1])
  unindented <- min(indents)
  indented<- max(indents)
  for (r in 1:nrow( ft$body[[8]]$pars$padding.left$data)){
    if (ft$body[[8]]$pars$padding.left$data[[r,1]]==indented){
      ft$body[[8]]$pars$keep_with_next$data[[r,1]]<-TRUE
      if(r>1){
        ft$body[[8]]$pars$keep_with_next$data[[r-1,1]]<-TRUE
      }
    } else {
      ft$body[[8]]$pars$keep_with_next$data[[r,1]]<-FALSE
      if(r>1){
        ft$body[[8]]$pars$keep_with_next$data[[r-1,1]]<-FALSE
      }
    }
  }
  return(ft)
}

gt_tbl_summary <- function(df_name,args){
  df<-get(df_name)
  vars_list<-args[[1]]
  # print(lapply(vars_list,make_match))
  by_var<-args[[2]]
  categorical<-args[[3]]
  continuous<-args[[4]]
  list_cont<-strsplit(continuous, "\\|")[[1]]
  summary_type_continuous <- args[[5]]
  
  ft <- df |>
  select(all_of(vars_list)) |>
  # https://rdrr.io/cran/gtsummary/man/tbl_summary.html for 2 rows check this out
  tbl_summary(
    by = {{ by_var }},
    type = all_continuous() ~ summary_type_continuous,
    statistic = list(
      all_continuous() ~ list_cont,
      all_categorical() ~ categorical
    ),
    missing = "no",
    label= lapply(vars_list,make_match)
  ) |>
  add_p(pvalue_fun = ~ style_pvalue(.x, digits = 2)) |>
  modify_caption("Patient Characteristics (N = {N})") |>
  modify_spanning_header(all_stat_cols() ~ match_single(by_var)) |>
  as_flex_table() |>
  keep_with_next()
  # print(gt)
  return(ft)
}


## -----------------------------------------------------------------------------
ggplot_bar <- function(df_name,args) {
  # arg 3 = df name
  # df<-get(args[[3]])
  df<-get(df_name)
  x_var <-ept(args[[2]])
  y_var <- ept(args[[1]])
  
  bar_graph<-ggplot(df, aes(x = factor(.data[[x_var]]), y = .data[[y_var]])) + 
    geom_bar(stat = "identity", fill = "grey", width = 0.5) +
    geom_text(aes(label = paste0(round(.data[[y_var]]*100,1),'%')), vjust = -0.5, color = "black", size = 3) +  
    labs(x = "evt_days", y = "Percentage") +
    ggtitle("Bar Plot of Percentages with Values") +
    scale_y_continuous(labels = scales::percent_format(accuracy = 1)) +
    theme_minimal()
    
  return(bar_graph)

}


## -----------------------------------------------------------------------------
aggregate_props <- function(df_name,args) {
    df<-get(df_name)
    base_df <-get(args[[3]])
    variable_names <- eval(parse(text=args[[1]]))
    freq_var <- eval(parse(text=args[[2]]))
    renamed_var <- args[[4]]
    print(paste('renamed_var',renamed_var))
    prop_dfs <- as_tibble(table(df[[freq_var]]),.name_repair = "minimal")
    names(prop_dfs)[1]<-freq_var
    is_factor <- is.factor(base_df[[freq_var]])
    if (!is_factor) {
      prop_dfs[[freq_var]]<- as.numeric(prop_dfs[[freq_var]])
    }
      
    for (var_name in variable_names) {
      prop_df <- aggregate(get(var_name) ~ get(freq_var), df, function(x) {
        prop <- prop.table(table(x))
        if (length(prop) == 1) {
          prop <- c(prop, 0)  # If only one value, add proportion of the other value as 0
        }
        return(prop[2])
      })
      names(prop_df)[1]<-freq_var
      names(prop_df)[2]<-var_name
      prop_dfs <- merge(prop_dfs, prop_df, by = freq_var, all = TRUE,sort=FALSE)
      
    }
    if (is_factor) {
      prop_dfs[[freq_var]]<-factor(prop_dfs[[freq_var]],levels(df[[freq_var]]))
    }
  prop_dfs<- prop_dfs |> rename_with(~ renamed_var,1)  
  return(prop_dfs)

}


## -----------------------------------------------------------------------------
merge_props <- function(df_name,args,report_name) {
  tb_list<-strsplit(args[[1]], split = '|', fixed = TRUE)[[1]]
  outcome_names <- strsplit(args[[2]], split = '|', fixed = TRUE)[[1]]
  df <- ept(tb_list[[1]]) |> mutate(tb=outcome_names[[1]])
  for (i in 2:length(tb_list)) {
    df<-bind_rows(df, ept(tb_list[[i]])|> mutate(tb=outcome_names[[i]]))
  }
  write.csv(df,paste0(report_name,'.csv'))
  # print(names(df))
  if('n' %in% colnames(df)) {
    df<-select(df, -n)
  }
  return(df)
}


## -----------------------------------------------------------------------------
ggplot_bar_gp <- function(df_name,args) {
  df<-get(df_name)
  x_var <-ept(args[[2]])
  y_var <- ept(args[[1]])
  
 bar_graph<-ggplot(data = df_merge,
       aes(x = .data[[x_var]], y = .data[[y_var]], group = factor(.data[['tb']]))
    ) +
    geom_bar(
        stat = "identity",
        aes(fill = factor(.data[['tb']])),
        position = position_dodge(width = 0.9)
    )
  return(bar_graph)
}


## -----------------------------------------------------------------------------
analysis <- read_excel('Analysis/recode.xlsx',sheet = "analysis")
analysis <- analysis |> filter(include!=0)
analysis


## -----------------------------------------------------------------------------

pattern0 <- str_c("\\b(?:", shf_2[,'var'], ")[\\$\\#]\\w+\\b", collapse = "|")
pattern1 <- "[\\$\\#]\\w+\\b"
pattern2 <- str_c("\\b",shf_1[,'var'],"\\b",collapse = "|")
patterns = c(pattern0, pattern1, pattern2)
analysis <- read_excel('Analysis/recode.xlsx',sheet = "analysis")
analysis <- analysis |> filter(include!=0)
report_name <- paste0("analysis/output/",report_file())
doc <- read_docx()
disp_ <- TRUE
for (i in 1:nrow(analysis)){
  switch(analysis[[i,'command']],
    "load"={
      processed_args<-get_args(recode,analysis[i,],0,1,patterns)
      data_loaded<-load_file(processed_args[[1]])
    },
    "rename"={
      processed_args<-get_args(recode,analysis[i,],0,1,patterns)
      original<- read_excel('Analysis/recode.xlsx',sheet = "original") 
      recode <- read_excel('Analysis/recode.xlsx',sheet = "recode")
      result<-rename_variables(data_loaded,recode,original,processed_args)
      recode<-result$recode
      sub_pri<-result$sub_pri
    },
    "recode"={
      result<-reload_recode(sub_pri,recode,original,patterns)
      df <- result$df
      recode <- result$recode
      doc<-recode_list(recode,analysis,i,doc)
      df<-factorize_variables(df,recode)
      df<-relevel_variables(df,recode)
    },
    "df" = {
      processed_args<-get_args(recode,analysis[i,],0,1,patterns)
      assign(analysis[[i,'out']],get(analysis[[i,'in']]))
      doc <- add_n(analysis[[i,'out']],doc)
    },
    "refilter" = {
      processed_args<-get_args(recode,analysis[i,],1,0,patterns,do_eval = FALSE,quote_var = FALSE)
      assign(analysis[[i,'out']], refilter(analysis[[i,'in']],processed_args)) 
      doc <- add_n(analysis[[i,'out']],doc,paste("filter: ",processed_args[1]))
    },
    "subset" = {
      processed_args<-get_args(recode,analysis[i,],1,0,patterns,do_eval = TRUE,quote_var = TRUE)[[1]]
      df_name_in <- analysis[[i,'in']]
      df_name_out <- analysis[[i,'out']]
      while  (nrow(analysis) > i && analysis[i,'cont']==1){
        i<-i+1
        processed_args<-c(processed_args,get_args(recode,analysis[i,],1,0,patterns,do_eval = TRUE,quote_var = TRUE)[[1]])
      }
      assign(df_name_out, subset(df_name_in,processed_args)) 
      doc <- add_n(analysis[[i,'out']],doc,paste("subset: ",processed_args[1]))
    },
    "ctn1" = {doc<-ctn1_shell(recode,analysis[i,],1,1,patterns,doc)},
    "ctn2" = {doc<-ctn2_shell(recode,analysis[i,],1,1,patterns,doc)},
    "ctn3" = {doc<-ctn3_shell(recode,analysis[i,],1,1,patterns,doc)},
    "ctn4" = {doc<-ctn4_shell(recode,analysis[i,],1,1,patterns,doc)},
    "summ" = {doc<-summ_shell(recode,analysis[i,],1,1,patterns,doc)},
    "freq2" = {doc<-freq2_shell(recode,analysis[i,],1,1,patterns,doc)},
    "freq1" = {doc<-freq1_shell(recode,analysis[i,],1,1,patterns,doc)},
    "glm"= {doc<-gt_glm_shell(recode,analysis[i,],1,0,patterns,doc)},
    "tbl_summary"= {
      doc<-gt_tbl_summary_shell(recode,analysis[i,],1,0,patterns,doc)
      },
    "aggregate_count" = {
      processed_args<-get_args(recode,analysis[i,],1,1,patterns,do_eval = FALSE,quote_var = FALSE)
      assign(analysis[[i,'out']],aggregate_count(analysis[[i,'in']],processed_args))
    },
    "aggregate_props" = {
      processed_args<-get_args(recode,analysis[i,],2,2,patterns,do_eval = FALSE,quote_var = TRUE)
      assign(analysis[[i,'out']],aggregate_props(analysis[[i,'in']],processed_args))
    },
    "merge_props" = {
      processed_args<-get_args(recode,analysis[i,],2,2,patterns,do_eval = FALSE,quote_var = TRUE)
      assign(analysis[[i,'out']],merge_props(analysis[[i,'in']],processed_args,report_name))
      doc <- doc |> body_add_break() |> add_text(paste0('Merged data:',report_name,'.csv saved in reports folder '),level=9) 
    },
    "ggplot_bar" = {
      processed_args<-get_args(recode,analysis[i,],2,1,patterns,do_eval = FALSE,quote_var = TRUE)
      bar_graph<-ggplot_bar(analysis[[i,'in']],processed_args)
      doc <- doc |> body_add_break() |> add_text( analysis[[i,'title']]) |>  body_add_gg(bar_graph, width = 6, height = 4)
    },
    "ggplot_bar_gp" = {
      processed_args<-get_args(recode,analysis[i,],2,1,patterns,do_eval = FALSE,quote_var = TRUE)
      bar_graph<-ggplot_bar_gp(analysis[[i,'in']],processed_args)
      doc <- doc |> body_add_break() |> add_text( analysis[[i,'title']]) |>  body_add_gg(bar_graph, width = 6, height = 4)
    },
    "ggplot_hist" = {
      processed_args<-get_args(recode,analysis[i,],1,1,patterns,do_eval = FALSE,quote_var = TRUE)
      plot<-ggplot_hist(analysis[[i,'in']],processed_args)
      doc <- doc |> body_add_break() |>add_text( analysis[[i,'title']]) |>body_add_gg(plot, width = 6, height = 4)
    },
    "ggplot_hist_l" = {
      processed_args<-get_args(recode,analysis[i,],1,1,patterns,do_eval = FALSE,quote_var = TRUE)
      doc<-ggplot_hist_l(analysis[[i,'in']],processed_args,doc,analysis[[i,'title']])
    },
    "title"={      doc <- doc |> body_add_break() |> add_text( analysis[[i,'title']],level =1)     }
)
}
file_name <- paste0(report_name,".docx")
print(file_name)
doc<-print(doc, target = file_name)  

