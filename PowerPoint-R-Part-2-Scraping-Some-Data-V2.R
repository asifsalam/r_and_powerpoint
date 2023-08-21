# Author: Asif Salam
# email: asif.salam@yahoo.com
# Date: 2023-08-06

# library(RCurl)
# library(XML)
library(tidyverse)
library(rvest)

data_dir <- "./data"
clint_eastwood_films_file <- file.path(data_dir,"clint_eastwood_films.tsv")
poster_dir <- "./posters"

actor_name <- "Clint Eastwood"
actor_url <- "http://www.imdb.com/name/nm0000142/?ref_=fn_al_nm_1"
local_file <- ".//imdb-clint2//Clint Eastwood - IMDb.htm"
local_page <- read_html(local_file)
# html_page <- read_html(actor_url)

filmography_div_selector <- "#accordion-item-actor-previous-projects"
filmography_div <- local_page %>% html_nodes(filmography_div_selector)
filmography_list <- filmography_div %>% html_elements("li.ipc-metadata-list-summary-item")

poster_list <- filmography_list %>% html_element("img")
# The src is actually the local path to the saved file, in this case
poster_urls <- poster_list %>% html_attr("src")

film_title_list <- filmography_list %>% html_element("div.ipc-metadata-list-summary-item__tc") %>% html_element("a.ipc-metadata-list-summary-item__t")
film_titles <- film_title_list %>% html_text2()
film_urls <- film_title_list %>% html_attr("href")

character_list <- filmography_list %>% html_element("div.ipc-metadata-list-summary-item__tc") %>% html_element("span.ipc-metadata-list-summary-item__li")
characters <- character_list %>% html_text2()

film_year_list <- filmography_list %>% html_element("div.ipc-metadata-list-summary-item__cc") %>% html_element("span.ipc-metadata-list-summary-item__li")
film_years <- film_year_list %>% html_text2()

additional_info_list <- filmography_list %>% html_element("div.ipc-metadata-list-summary-item__tc")
# create a column for additional information that can be used for filtering out data
additional_info_text <- additional_info_list %>% html_text2() %>% str_replace_all("\\n",";")

film_id = paste0("CE",str_pad(seq(1,length(film_titles)),2,side="left",pad="0"))
film_key = str_to_lower(gsub("[[:punct:] ]+","",film_titles))
# poster_file = paste0("posters//poster",film_id,".jpg"))
poster_file = file.path(poster.dir,paste0("poster",film_id,".jpg"))

# download.file(poster_list_url[1],poster_file[1], mode="wb")
# Since the file is local, we will use file.copy

file.copy(paste0("imdb-clint/",str_replace_all(poster_urls,"%20"," ")),poster_file)

clint_eastwood_filmography <- tibble(key=film_key,id=film_id,title=film_titles,film_url = film_urls,
             character = characters, release_year = film_years, poster_url = poster_urls,
             poster_file = poster_file, additional_info = additional_info_text)

# Save the data frame
write.table(clint_eastwood_filmography,file="./data/clint_eastwood_films.tsv",append=FALSE,quote=TRUE,sep="\t",row.names=FALSE)


