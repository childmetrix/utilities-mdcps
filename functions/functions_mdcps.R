#####################################
#####################################
# FUNCTIONS SPECIFIC TO MDCPS 
#####################################
#####################################

#####################################
# Insert column holding quarter, make it first column ----
#####################################

# Usage
# df <- some_df %>%
#  add_quarter_col

# Insert quarter column and make it the first column (e.g., "2025 Q1")
add_quarter_col <- function(df) {
  df %>%
    mutate(quarter = folder_date_quarter, .before = 1)
}

#####################################
# Create quarterly and cumulative workbooks ----
#####################################

# Saves quarterly workbook to a dated run folder and a cumulative file
# (up to the current quarter) to the cumulative root (no subfolder).
#
# Args:
#   sheets_to_save           named list of data.frames (sheet_name = df)
#   base_data_dir            "D:/repo_mdcps_suspension_period/r_2.9.a/data"
#   folder_date              e.g., "2025_Q2"
#   commitment               e.g., "2.9.a"
#   commitment_description   e.g., "MIC Manual Review Verification"
#   folder_cumulative        e.g., file.path(base_data_dir, "2025_cumulative")
#   create_latest_copies     optional TRUE to also drop a 'latest' convenience copy
#
# Returns: list(quarterly_xlsx = <path>, cumulative_xlsx = <path or NA>)
# ------------------------------------------------------------

# ------------------------------------------------------------
# run_quarter_and_cumulative()  — preserves POSIXct datetimes
# ------------------------------------------------------------
run_quarter_and_cumulative <- function(
    sheets_to_save,
    base_data_dir,
    folder_date,                # e.g., "2025_Q2"
    commitment,                 # e.g., "2.9.a"
    commitment_description,     # e.g., "MIC Manual Review Verification"
    folder_cumulative,          # e.g., file.path(base_data_dir, "2025_cumulative")
    create_latest_copies = FALSE,
    # Optional hints (by column name) to force types when reading cumulative inputs
    datetime_cols_hint = c("intake_datetime","screening_datetime","initiation_datetime"),
    date_cols_hint     = character(0)
) {
  # ---------- local helpers ----------
  today_str <- format(Sys.Date(), "%Y-%m-%d")
  EXCEL_ORIGIN_DATE <- as.Date("1899-12-30")
  EXCEL_ORIGIN_POSIX <- as.POSIXct("1899-12-30", tz = "UTC")
  
  safe_label <- function(x) {
    x <- gsub("\\s*-\\s*-\\s*", " - ", x)
    x <- gsub("\\s{2,}", " ", x)
    x <- gsub("^\\s*-\\s*|\\s*-\\s*$", "", x)
    x
  }
  
  make_run_folder <- function(root_dir) {
    run_dir <- file.path(root_dir, today_str)
    if (!dir.exists(run_dir)) dir.create(run_dir, recursive = TRUE)
    writeLines(today_str, file.path(root_dir, "latest_run.txt"))
    run_dir
  }
  
  get_latest_run <- function(root_dir) {
    p <- file.path(root_dir, "latest_run.txt")
    if (file.exists(p)) {
      v <- readLines(p, warn = FALSE)
      if (length(v)) return(file.path(root_dir, v[1]))
    }
    subs <- list.dirs(root_dir, full.names = FALSE, recursive = FALSE)
    subs <- subs[grepl("^\\d{4}-\\d{2}-\\d{2}$", subs)]
    if (!length(subs)) return(NA_character_)
    file.path(root_dir, sort(subs, decreasing = TRUE)[1])
  }
  
  save_latest_copy <- function(file_path, root_dir, link_name = NULL) {
    latest_dir <- file.path(root_dir, "latest")
    if (!dir.exists(latest_dir)) dir.create(latest_dir, recursive = TRUE)
    if (is.null(link_name)) link_name <- basename(file_path)
    file.copy(file_path, file.path(latest_dir, link_name), overwrite = TRUE)
  }
  
  parse_folder_date <- function(folder_date) {
    m <- regexec("^(\\d{4})_Q([1-4])$", folder_date)
    r <- regmatches(folder_date, m)[[1]]
    if (length(r) != 3) stop("folder_date must look like 'YYYY_Q#'")
    list(year4 = r[2], q = as.integer(r[3]))
  }
  
  # ---------- type normalization (Date vs POSIXct) ----------
  # Decide which columns are datetime/date by name + current classes
  classify_cols <- function(df, datetime_hint, date_hint) {
    nms <- names(df)
    # infer datetime by name pattern
    name_is_dt <- grepl("_datetime$", nms, ignore.case = TRUE) | nms %in% datetime_hint
    # infer from class present
    class_is_dt <- vapply(df, function(x) inherits(x, c("POSIXct","POSIXt")), logical(1))
    class_is_date <- vapply(df, function(x) inherits(x, "Date"), logical(1))
    # user hint for dates
    name_is_date <- nms %in% date_hint
    
    datetime_cols <- nms[name_is_dt | class_is_dt]
    # date columns are those hinted or already Date, but not in datetime
    date_cols <- setdiff(nms[name_is_date | class_is_date], datetime_cols)
    
    list(datetime = unique(datetime_cols), date = unique(date_cols))
  }
  
  # Convert typed/character/numeric to Date or POSIXct based on target
  to_date <- function(x) {
    if (inherits(x, "Date")) return(x)
    if (inherits(x, c("POSIXct","POSIXt"))) return(as.Date(x))
    if (is.numeric(x)) {
      # Excel serial days
      return(as.Date(x, origin = EXCEL_ORIGIN_DATE))
    }
    if (is.character(x)) {
      # 1) Excel serial provided as text
      if (grepl("^\\s*\\d+(\\.\\d+)?\\s*$", x)) {
        return(as.Date(as.numeric(x), origin = EXCEL_ORIGIN_DATE))
      }
      # 2) ISO then US fallback, while PRESERVING Date class (no ifelse())
      y1 <- suppressWarnings(as.Date(x))                     # ISO (YYYY-mm-dd)
      y2 <- suppressWarnings(as.Date(x, format = "%m/%d/%Y"))# US  (mm/dd/YYYY)
      y  <- y1
      nas <- is.na(y1)
      if (any(nas)) y[nas] <- y2[nas]
      return(y)
    }
    x
  }
  
  to_datetime <- function(x, tz = "UTC") {
    if (inherits(x, c("POSIXct","POSIXt"))) return(x)
    if (inherits(x, "Date")) return(as.POSIXct(x, tz = tz))  # midnight
    if (is.numeric(x)) {
      # Excel serial (fractional days)
      return(as.POSIXct(EXCEL_ORIGIN_POSIX + x * 86400, tz = tz))
    }
    if (is.character(x)) {
      # Excel serial as text
      if (grepl("^\\s*\\d+(\\.\\d+)?\\s*$", x)) {
        return(as.POSIXct(EXCEL_ORIGIN_POSIX + as.numeric(x) * 86400, tz = tz))
      }
      # ISO (space or T)
      y1 <- suppressWarnings(as.POSIXct(x, tz = tz))
      if (!all(is.na(y1))) return(y1)
      # Common US patterns
      y2 <- suppressWarnings(strptime(x, "%m/%d/%Y %H:%M:%S", tz = tz))
      if (!all(is.na(y2))) return(as.POSIXct(y2, tz = tz))
      y3 <- suppressWarnings(strptime(x, "%m/%d/%Y %H:%M", tz = tz))
      if (!all(is.na(y3))) return(as.POSIXct(y3, tz = tz))
      # Date-only fallback
      y4 <- suppressWarnings(as.Date(x))
      if (!all(is.na(y4))) return(as.POSIXct(y4, tz = tz))
    }
    x
  }
  
  normalize_types <- function(df, datetime_cols, date_cols, tz = "UTC") {
    for (nm in intersect(names(df), datetime_cols)) df[[nm]] <- to_datetime(df[[nm]], tz = tz)
    for (nm in setdiff(intersect(names(df), date_cols), datetime_cols)) df[[nm]] <- to_date(df[[nm]])
    df
  }
  
  # Write one sheet with explicit Excel styles for Date and POSIXct
  write_sheet_typed <- function(wb, sheet_name, df, datetime_hint, date_hint) {
    addWorksheet(wb, sheet_name)
    
    # detect columns by current classes + hints
    cls <- classify_cols(df, datetime_hint, date_hint)
    df2 <- normalize_types(df, cls$datetime, cls$date)
    
    # default Excel formats
    old_opt <- getOption("openxlsx.dateFormat")
    on.exit(options(openxlsx.dateFormat = old_opt), add = TRUE)
    options(openxlsx.dateFormat = "yyyy-mm-dd")
    
    writeData(wb, sheet = sheet_name, x = df2)
    
    # Apply styles
    if (nrow(df2) > 0) {
      # Date
      date_cols <- which(vapply(df2, inherits, logical(1), "Date"))
      if (length(date_cols)) {
        st_date <- createStyle(numFmt = "yyyy-mm-dd")
        addStyle(wb, sheet = sheet_name, style = st_date,
                 rows = 2:(nrow(df2) + 1), cols = date_cols,
                 gridExpand = TRUE, stack = TRUE)
      }
      # POSIXct
      dt_cols <- which(vapply(df2, function(x) inherits(x, c("POSIXct","POSIXt")), logical(1)))
      if (length(dt_cols)) {
        st_dt <- createStyle(numFmt = "yyyy-mm-dd hh:mm:ss")
        addStyle(wb, sheet = sheet_name, style = st_dt,
                 rows = 2:(nrow(df2) + 1), cols = dt_cols,
                 gridExpand = TRUE, stack = TRUE)
      }
    }
  }
  
  # Gather latest (non-cumulative) file for Q1..max_q only
  find_latest_quarterly_files_upto <- function(base_data_dir, year4, max_q, pattern_keep = NULL) {
    out <- character(0)
    for (q in 1:max_q) {
      q_name <- paste0(year4, "_Q", q)
      root <- file.path(base_data_dir, q_name, "processed")
      latest_run <- get_latest_run(root)
      if (is.na(latest_run)) next
      files <- list.files(latest_run, pattern = "\\.xlsx$", full.names = TRUE)
      files <- files[!grepl("Cumulative", basename(files), ignore.case = TRUE)]
      if (!is.null(pattern_keep)) files <- files[grepl(pattern_keep, basename(files), ignore.case = TRUE)]
      if (!length(files)) next
      files <- files[order(file.info(files)$mtime, decreasing = TRUE)]
      out <- c(out, files[1])
    }
    out
  }
  
  # Build cumulative from vector of quarterly paths (preserving datetimes)
  build_cumulative_from_files <- function(quarterly_paths, cumulative_xlsx_path, sheet_names,
                                          datetime_hint, date_hint) {
    wb_out <- createWorkbook()
    for (sheet in sheet_names) {
      dfs <- list()
      for (p in quarterly_paths) {
        ok <- TRUE
        # IMPORTANT: do not auto-convert; we’ll normalize ourselves
        df <- tryCatch(readWorkbook(p, sheet = sheet, detectDates = FALSE),
                       error = function(e) { ok <<- FALSE; NULL })
        if (ok && !is.null(df)) dfs[[length(dfs) + 1]] <- df
      }
      if (!length(dfs)) next
      common_cols <- Reduce(intersect, lapply(dfs, names))
      dfs <- lapply(dfs, function(x) x[, common_cols, drop = FALSE])
      df_all <- do.call(rbind, dfs)
      
      # Normalize by hints and heuristics (so datetimes keep time)
      classes <- classify_cols(df_all, datetime_hint, date_hint)
      df_all  <- normalize_types(df_all, classes$datetime, classes$date)
      
      write_sheet_typed(wb_out, sheet, df_all, datetime_hint, date_hint)
    }
    saveWorkbook(wb_out, cumulative_xlsx_path, overwrite = TRUE)
  }
  # ---------- end helpers ----------
  
  # Paths
  folder_quarter_root <- file.path(base_data_dir, folder_date, "processed")
  
  # 1) QUARTERLY (not cumulative): write each sheet with proper date/datetime handling
  quarter_run_dir <- make_run_folder(folder_quarter_root)
  quarterly_name  <- safe_label(paste(folder_date, commitment, commitment_description, today_str, sep = " - "))
  quarterly_xlsx  <- file.path(quarter_run_dir, paste0(quarterly_name, ".xlsx"))
  
  wb <- createWorkbook()
  for (sheet_name in names(sheets_to_save)) {
    write_sheet_typed(
      wb, sheet_name,
      df = sheets_to_save[[sheet_name]],
      datetime_hint = datetime_cols_hint,
      date_hint     = date_cols_hint
    )
  }
  saveWorkbook(wb, quarterly_xlsx, overwrite = TRUE)
  if (create_latest_copies) save_latest_copy(quarterly_xlsx, folder_quarter_root)
  
  # 2) CUMULATIVE (file only): include Q1..current Q only, preserving datetimes
  pd <- parse_folder_date(folder_date)
  year4 <- pd$year4
  max_q <- pd$q
  
  latest_quarterlies <- find_latest_quarterly_files_upto(base_data_dir, year4, max_q)
  
  if (length(latest_quarterlies)) {
    cumulative_file_stem <- safe_label(
      paste(folder_date, commitment, commitment_description, "Cumulative", today_str, sep = " - ")
    )
    cumulative_xlsx <- file.path(folder_cumulative, paste0(cumulative_file_stem, ".xlsx"))
    
    build_cumulative_from_files(
      quarterly_paths       = latest_quarterlies,
      cumulative_xlsx_path  = cumulative_xlsx,
      sheet_names           = names(sheets_to_save),
      datetime_hint         = datetime_cols_hint,
      date_hint             = date_cols_hint
    )
    
    writeLines(basename(cumulative_xlsx), file.path(folder_cumulative, "latest_cumulative.txt"))
    if (create_latest_copies) save_latest_copy(cumulative_xlsx, folder_cumulative, link_name = "Cumulative-latest.xlsx")
  } else {
    cumulative_xlsx <- NA_character_
  }
  
  list(
    quarterly_xlsx  = quarterly_xlsx,
    cumulative_xlsx = cumulative_xlsx
  )
}

# Usage
# -----------------------------------

# base_data_dir       <- "D:/repo_mdcps_suspension_period/r_2.9.a/data"
# folder_cumulative   <- file.path(base_data_dir, substr(folder_date, 1, 4), "2025_cumulative")  # or your existing path

# sheets_to_save <- list(
#   "Population Details"        = detail_df,
#   "Manual Review Children"    = mr_children_df,
#   "Manual Review Allegations" = mr_allegations_df
# )
# 
# out <- run_quarter_and_cumulative(
#   sheets_to_save           = sheets_to_save,
#   base_data_dir            = base_data_dir,
#   folder_date              = folder_date,              # e.g., "2025_Q2"
#   commitment               = commitment,               # e.g., "2.9.a"
#   commitment_description   = commitment_description,   # e.g., "MIC Manual Review Verification"
#   folder_cumulative        = folder_cumulative,
#   create_latest_copies     = FALSE
# )

########################################
# Load file from sharefile ----
########################################

#’ Load a CSV or XLSX from a ShareFile‐style folder hierarchy
#’
#’ This function searches under a top‐level “base_path” for a subfolder matching
#’ the global `folder_date` (in one of YYYY_CY, YYYY_Qn or YYYY_MM formats).  
#’ If found, it optionally descends into date‐stamped subfolders to pick the
#’ most recent, then loads the first file whose name contains `file_string`.
#’
#’ @param base_path     Character. Root directory under which to look for period folders.
#’ @param file_string   Character. Substring to match in the target filename (before “.csv” or “.xlsx”).
#’ @param df_name       Character. Name of the data‐frame to assign into the global environment.
#’ @param sheet_name    Character (optional). Excel sheet name to read; if `NULL`, reads the first sheet.
#’ @param fallback_Q4   Logical (default FALSE).  Only when `folder_date` is “YYYY_CY”:  
#’                       if no year‐folder is found and this is `TRUE`,  
#’                       the loader will fall back to looking for Q4 of that year.
#’
#’ @return Invisibly returns the path of the file that was loaded;  
#’         also assigns the data‐frame named `df_name` into `.GlobalEnv`.
#’
#’ @details  
#’ - Year‐folders recognized: `YYYY_CY`, `YYYY CY`, `YYYY_ALL`, or `CY_YYYY`.  
#’ - Quarter‐folders: any name containing “Q<1–4>” + year in either order.  
#’ - Month‐folders: recognized by numeric (01–12) or full month name, plus derived quarter.  
#’ - After locating the correct period folder, if it contains dated subfolders (`YYYY-MM-DD`),
#’   the most recent date is used.  
#’ - Supports both CSV and XLSX; for XLSX you can specify a sheet.
#’
#’ @examples
#’ \dontrun{
#’   # Standard quarter load:
#’   folder_date <- "2024_Q2"
#’   load_sharefile_file("S:/Projects", "summary", "df2")
#’
#’   # Year load, falling back to Q4 if no 2024_CY folder exists:
#’   folder_date <- "2024_CY"
#’   load_sharefile_file("S:/Projects", "full_data", "dfY", fallback_Q4 = TRUE)
#’ }

load_sharefile_file <- function(base_path,
                                file_string,
                                df_name,
                                sheet_name    = NULL,
                                fallback_Q4   = FALSE,
                                guess_max     = 5000) {
  parts <- strsplit(folder_date, "_")[[1]]
  if (length(parts) != 2) {
    stop("`folder_date` must be in 'YYYY_CY', 'YYYY_Qn', or 'YYYY_MM' form.")
  }
  year   <- parts[1]
  period <- parts[2]
  
  # Build the primary patterns
  if (grepl("^CY$", period, ignore.case = TRUE)) {
    # year-based
    patterns_year <- c(
      paste0("(?i)^", year, "_CY$"),
      paste0("(?i)^", year, " CY$"),
      paste0("(?i)^", year, "_ALL$"),
      paste0("(?i)^CY_", year, "$")
    )
    patterns <- patterns_year
    
  } else if (grepl("Q", period, ignore.case = TRUE)) {
    # quarter-based
    q <- period
    patterns <- c(
      paste0("(?i)", year, ".*", q),
      paste0("(?i)", q, ".*", year)
    )
    
  } else {
    # month-based
    m  <- as.integer(period)
    q4 <- paste0("Q", ceiling(m / 3))
    mn <- month.name[m]
    patterns <- c(
      paste0("(?i)", year, ".*", q4),
      paste0("(?i)", q4, ".*", year),
      paste0("(?i)", year, ".*", sprintf("%02d", m)),
      paste0("(?i)", sprintf("%02d", m), ".*", year),
      paste0("(?i)", year, ".*", mn),
      paste0("(?i)", mn, ".*", year)
    )
  }
  
  # helper to find first matching subdir
  find_dir <- function(pats) {
    subdirs <- list.dirs(base_path, full.names = TRUE, recursive = FALSE)
    for (d in subdirs) {
      nm <- basename(d)
      if (any(vapply(pats, grepl, logical(1), x = nm, perl = TRUE))) {
        return(d)
      }
    }
    NULL
  }
  
  # 1) try primary patterns
  target_folder <- find_dir(patterns)
  
  # 2) if year-based + fallback requested + nothing found → try Q4 patterns
  if (is.null(target_folder) &&
      grepl("^CY$", period, ignore.case = TRUE) &&
      fallback_Q4) {
    message("  • CY folder not found; falling back to Q4 folder for ", year)
    q4_patterns <- c(
      paste0("(?i)Q4.*", year),
      paste0("(?i)", year, ".*Q4")
    )
    target_folder <- find_dir(q4_patterns)
  }
  
  if (is.null(target_folder)) {
    stop("No subfolder matching `", folder_date, "` found in ", base_path)
  }
  
  # If that folder contains date‐named subfolders, pick the most recent
  subsub <- list.dirs(path = target_folder, full.names = TRUE, recursive = FALSE)
  date_pattern <- "^\\d{4}-\\d{2}-\\d{2}$"
  dated_dirs   <- subsub[grepl(date_pattern, basename(subsub))]
  if (length(dated_dirs) > 0) {
    dates <- as.Date(basename(dated_dirs))
    target_folder <- dated_dirs[which.max(dates)]
  }
  
  # Look for the file
  # — if the user supplied an anchored regex (ends with $), use it verbatim;
  #   otherwise wrap it so we still only grab .csv/.xlsx files
  # e.g., file_string <- "_quarterly\\.xlsx$"
  if (grepl("\\$$", file_string)) {
    file_pattern <- file_string
  } else {
    file_pattern <- paste0(".*", file_string, ".*\\.(csv|xlsx)$")
  }
  found <- list.files(
    path       = target_folder,
    pattern    = file_pattern,
    full.names = TRUE
  )
  # file_pattern <- paste0(".*", file_string, ".*\\.(csv|xlsx)$")
  # found <- list.files(path = target_folder, pattern = file_pattern, full.names = TRUE)
  if (length(found) == 0) {
    stop("No file containing '", file_string, "' found in ", target_folder)
  }
  target_file <- found[1]
  
  # Read it in
  ext <- tolower(tools::file_ext(target_file))
  if (ext == "csv") {
    df <- readr::read_csv(target_file)
    
  } else if (ext == "xlsx") {
    # build the list of args so we only add sheet if needed
    args <- list(path      = target_file,
                 guess_max = guess_max)
    if (!is.null(sheet_name)) args$sheet <- sheet_name
    
    df <- do.call(readxl::read_excel, args)
    
  } else {
    stop("Unsupported file extension: ", ext)
  }
  
  # Assign to the global env
  assign(df_name, df, envir = .GlobalEnv)
}

########################################
# Clean custody file ----
########################################

# Description
# --------------------------------------

# Returns a cleaned custody_df:
# - include Q1 fix for a specific child
# - for children with multiple custody episodes, removes them if dates overlap
# - for children with duplicate custody episodes (same child_id, custody_start),
# keeps the one with lowest custody_sequence_number

# Usage - Option 1
# --------------------------------------

# custody_df <- process_custody(custody_df, folder_date_quarter)

# Usage - Option 2 (also returns the diagnostic frames that hold excluded, flagged cases)
# --------------------------------------

# res <- process_custody(custody_df, folder_date_quarter, return_details = TRUE)
# list2env(res, .GlobalEnv) # careful: overwrites objects with the same names

# Function
# --------------------------------------

process_custody <- function(custody_df, folder_date_quarter, return_details = FALSE) {
  
  far_future <- as.Date("9999-12-31")
  
  strict_overlap_pairs <- function(df, id, start, end) {
    x <- df %>%
      transmute(
        child_id,
        id    = .data[[id]],
        start = .data[[start]],
        end   = coalesce(.data[[end]], far_future)
      )
    
    x %>%
      inner_join(x, by = "child_id", suffix = c("_a", "_b"), relationship = "many-to-many") %>%
      filter(id_a != id_b) %>%
      # strict overlap (touching at boundary is OK)
      filter(start_a < end_b, start_b < end_a) %>%
      transmute(
        child_id,
        id_lo = pmin(id_a, id_b),
        id_hi = pmax(id_a, id_b)
      ) %>%
      distinct()
  }
  
  # 1) Errant Q1 fix (only if applicable)
  if (isTRUE(folder_date_quarter == "2025 Q1")) {
    custody_df <- custody_df %>%
      filter(!(child_id == 5823501 & custody_sequence_number == 4)) %>%
      mutate(
        fix_row = (child_id == 5823501 & custody_sequence_number == 3),
        isn                     = if_else(fix_row, 69753L, isn),
        custody_start           = if_else(fix_row, as.Date("2024-04-24"), custody_start),
        date_of_birth           = if_else(fix_row, as.Date("2007-03-06"), date_of_birth),
        custody_tracking_number = if_else(fix_row, "20YC24P00571", custody_tracking_number),
        service_type            = if_else(fix_row, "Placement Services - R & S", service_type),
        service_outcome         = if_else(fix_row, "Living with other relatives", service_outcome)
      ) %>%
      select(-fix_row)
  }
  
  # 2) Remove strict overlaps within child_id
  custody_pairs <- strict_overlap_pairs(
    custody_df,
    id    = "custody_sequence_number",
    start = "custody_start",
    end   = "custody_end"
  )
  
  # participants as (child_id, id) -> join using id as custody_sequence_number
  custody_overlap_participants <- bind_rows(
    custody_pairs %>% transmute(child_id, id = id_lo),
    custody_pairs %>% transmute(child_id, id = id_hi)
  ) %>% distinct()
  
  custody_overlaps_df <- custody_df %>%
    semi_join(custody_overlap_participants,
              by = c("child_id", "custody_sequence_number" = "id"))
  
  custody_df_clean <- custody_df %>%
    anti_join(custody_overlap_participants,
              by = c("child_id", "custody_sequence_number" = "id"))
  
  # 3) Remove duplicates (keep lowest custody_sequence_number per child_id + start)
  duplicates_custody_df <- custody_df_clean %>%
    group_by(child_id, custody_start) %>%
    filter(n() > 1) %>%
    mutate(kept = suppressWarnings(custody_sequence_number == min(custody_sequence_number, na.rm = TRUE))) %>%
    ungroup()
  
  custody_df_clean <- custody_df_clean %>%
    group_by(child_id, custody_start) %>%
    slice_min(custody_sequence_number, with_ties = FALSE) %>%
    ungroup()
  
  # Return
  if (isTRUE(return_details)) {
    return(list(
      custody_df             = custody_df_clean,
      custody_overlaps_df    = custody_overlaps_df,
      duplicates_custody_df  = duplicates_custody_df
    ))
  } else {
    return(custody_df_clean)
  }
}

########################################
# Clean placements file ----
########################################

# Description
# --------------------------------------

# Returns a cleaned placements_df:

# - for children with multiple placement episodes, removes them if dates overlap
# - joins to custody_df (the cleaned, processed version) and removes placements 
# whose dates fall outside of the custody dates, with a 1-day grace before 
# custody start
# - for children with duplicate placements (same child_id, placement_id) but 
# different custody_sequence_number, keeps the one with lowest 
# custody_sequence_number
# - for children with remaining duplicate placements (same child_id, placement_id), 
# keeps only one (doesn't matter which). These are quasi-duplicates due to 
# having 2 or more reasons for emergency placement

# Usage - Option 1
# --------------------------------------

# placements_df <- process_placements(placements_df, custody_df)

# Usage - Option 2 (also returns the diagnostic frames that hold excluded, flagged cases)
# --------------------------------------

# res <- process_placements(placements_df, custody_df, return_details = TRUE)
# list2env(res, .GlobalEnv) # careful: overwrites objects with the same names

# Function
# --------------------------------------

process_placements <- function(placements_df, custody_df, grace_days = 1, return_details = FALSE) {
  
  # --- quick schema checks
  need_custody <- c("child_id","custody_sequence_number","custody_start","custody_end")
  if (!all(need_custody %in% names(custody_df))) {
    stop("custody_df is missing required columns: ",
         paste(setdiff(need_custody, names(custody_df)), collapse = ", "))
  }
  need_place <- c("child_id","placement_id","custody_sequence_number","placement_start","placement_end")
  if (!all(need_place %in% names(placements_df))) {
    stop("placements_df is missing required columns: ",
         paste(setdiff(need_place, names(placements_df)), collapse = ", "))
  }
  
  far_future <- as.Date("9999-12-31")
  
  strict_overlap_pairs <- function(df, id, start, end) {
    x <- df %>%
      transmute(
        child_id,
        id    = .data[[id]],
        start = .data[[start]],
        end   = coalesce(.data[[end]], far_future)
      )

    x %>%
      inner_join(x, by = "child_id", suffix = c("_a", "_b"), relationship = "many-to-many") %>%
      filter(id_a != id_b) %>%                         # different placements
      # strict overlap (touching exactly at a boundary is OK)
      filter(start_a < end_b, start_b < end_a) %>%
      transmute(
        child_id,
        id_lo = pmin(id_a, id_b),
        id_hi = pmax(id_a, id_b)
      ) %>%
      distinct()
  }
  
  participants_from_pairs <- function(pairs, id_col_out) {
    bind_rows(
      pairs %>% transmute(child_id, !!id_col_out := id_lo),
      pairs %>% transmute(child_id, !!id_col_out := id_hi)
    ) %>% distinct()
  }
  
  # --- placements with strict overlaps (within child_id)
  placement_pairs <- strict_overlap_pairs(
    placements_df,
    id    = "placement_id",
    start = "placement_start",
    end   = "placement_end"
  )
  
  placement_overlap_participants <-
    participants_from_pairs(placement_pairs, id_col_out = "placement_id")
  
  placements_overlaps_df <- placements_df %>%
    semi_join(placement_overlap_participants, by = c("child_id","placement_id"))
  
  placements_df_clean <- placements_df %>%
    anti_join(placement_overlap_participants, by = c("child_id","placement_id"))
  
  cat("Placement records removed due to overlap:",
      nrow(placements_df) - nrow(placements_df_clean), "\n")
  
  # --- placements outside custody episodes (1-day grace before custody_start)
  required_cols <- c(
    "placement_id","child_id","custody_sequence_number","custody_start","custody_end",
    "placement_start","placement_end","facility_type","facility_group"
  )
  
  placements_df_clean <- placements_df_clean %>%
    left_join(
      custody_df %>% select(child_id, custody_sequence_number, custody_start, custody_end),
      by = c("child_id","custody_sequence_number")
    ) %>%
    mutate(
      custody_end2   = coalesce(custody_end,   far_future),
      placement_end2 = coalesce(placement_end, far_future),
      
      # allow same-day or up to 'grace_days' BEFORE custody_start
      start_before_custody_start = !is.na(placement_start) & !is.na(custody_start) &
        placement_start < (custody_start - grace_days),
      
      start_after_custody_end  = !is.na(placement_start) & placement_start > custody_end2,
      end_before_custody_start = !is.na(placement_end)   & placement_end   < custody_start,
      end_after_custody_end    = placement_end2 > custody_end2,  # covers open placements too
      
      outside_any = start_before_custody_start |
        start_after_custody_end    |
        end_before_custody_start   |
        end_after_custody_end
    ) %>%
    select(placement_id, child_id, custody_sequence_number, custody_start, custody_end, everything())
  
  placements_outside_df <- placements_df_clean %>%
    filter(outside_any) %>%
    arrange(child_id, custody_sequence_number, placement_start) %>%
    select(any_of(required_cols),
           start_before_custody_start, start_after_custody_end,
           end_before_custody_start, end_after_custody_end, outside_any)
  
  cat("Number of unique children with placements outside custody episodes:",
      n_distinct(placements_outside_df$child_id), "\n")
  
  placements_df_clean <- placements_df_clean %>%
    filter(!outside_any)
  
  # --- duplicates (same child_id + placement_id across multiple custody episodes)
  duplicates_placements_df <- placements_df_clean %>%
    group_by(child_id, placement_id) %>%
    filter(n_distinct(custody_sequence_number) > 1) %>%
    mutate(kept = suppressWarnings(custody_sequence_number == min(custody_sequence_number, na.rm = TRUE))) %>%
    ungroup()
  
  cat("Number of duplicate placements found:", nrow(duplicates_placements_df), "\n")
  
  placements_df_clean <- placements_df_clean %>%
    group_by(child_id, placement_id) %>%
    filter(
      custody_sequence_number == min(custody_sequence_number) |
        n_distinct(custody_sequence_number) == 1
    ) %>%
    ungroup()
  
  # --- quasi duplicates (multiple rows per placement due to reasons, etc.)
  remaining_duplicates_df <- placements_df_clean %>%
    group_by(child_id, placement_id) %>%
    filter(n() > 1) %>%
    ungroup() %>%
    select(
      child_id,
      custody_sequence_number,
      custody_start,
      custody_end,
      placement_id,
      placement_sequence_per_episode,
      placement_start,
      placement_end,
      facility_type,
      cong_care_reasons,
      emer_plac_reasons
    ) %>%
    arrange(child_id, placement_id)
  
  cat("Number of quasi duplicate placements found:", nrow(remaining_duplicates_df), "\n")
  
  # Keep one row per (child_id, placement_id); exact record kept doesn't matter
  placements_df_clean <- placements_df_clean %>%
    distinct(child_id, placement_id, .keep_all = TRUE)
  
  if (isTRUE(return_details)) {
    return(list(
      placements_df              = placements_df_clean,
      placements_overlaps_df     = placements_overlaps_df,
      placements_outside_df      = placements_outside_df,
      duplicates_placements_df   = duplicates_placements_df,
      remaining_duplicates_df    = remaining_duplicates_df
    ))
  } else {
    return(placements_df_clean)
  }
}

########################################
# Get month review period in YYYY-MM MMM form ----
########################################

# Used?

# Used for 3.1 non relative, to filter Eric's detail sheet according to the month
# of interest. His values are 2024-01 Jan, 2024-02 Feb, etc.
get_review_period <- function() {
  # Convert reporting_period_start to a Date object
  date_obj <- as.Date(reporting_period_start)
  
  # Extract year, month number, and abbreviated month name
  year <- format(date_obj, "%Y")
  month_num <- format(date_obj, "%m")
  month_abbrev <- format(date_obj, "%b")
  
  # Construct the review period string
  review_period <- paste0(year, "-", month_num, " ", month_abbrev)
  return(review_period)
}


########################################
# 3.1 non-relative ----
########################################

# Function to find memo for the given month and extract:
# 1. # of homes licensed and overdue. (From this we'll calculate % and hope it 
# matches what they put in the PDF.)
# 2. List of resource IDs for homes re-opened or converted (we send these to PC
# for their review and verification they were done according to policy). The
# regex was really fun for this. :)

get_mdcps_3_1_counts <- function(base_path, folder_date,
                                 pattern_pdf  = "a and b memo",
                                 pattern_docx = "a and b memo") {
  
  # Convert "YYYY_MM" -> "Month YYYY" if needed
  to_month_yyyy <- function(x) {
    dt <- as.Date(paste0(x, "_01"), format = "%Y_%m_%d")
    paste0(month.name[as.integer(format(dt, "%m"))], " ", format(dt, "%Y"))
  }
  
  # Start-of-month Date from folder_date in either "YYYY_MM" or "Month YYYY"
  month_start_date <- function(x) {
    if (grepl("^\\d{4}_[01]\\d$", x)) {
      as.Date(paste0(x, "_01"), format = "%Y_%m_%d")
    } else {
      as.Date(paste0("01 ", x), format = "%d %B %Y")
    }
  }
  
  # Normalize text quirks from PDFs/DOCX
  normalize_text <- function(txt) {
    txt <- gsub("\u00A0", " ", txt, fixed = TRUE)                               # NBSP -> space
    txt <- gsub("([A-Za-z])-[[:space:]]+([A-Za-z])", "\\1\\2", txt, perl = TRUE) # undo hyphen linebreaks
    txt <- gsub("[[:space:]]+", " ", txt)
    trimws(txt)
  }
  
  # Choose subdir: either folder_date or "Month YYYY"
  candidates <- unique(c(
    file.path(base_path, folder_date),
    if (grepl("^\\d{4}_[01]\\d$", folder_date)) file.path(base_path, to_month_yyyy(folder_date)) else character(0)
  ))
  existing <- candidates[dir.exists(candidates)]
  if (!length(existing)) stop("No subdirectory found for '", folder_date, "' under: ", base_path)
  subdir <- existing[1]
  
  report_month <- month_start_date(folder_date)
  month_label  <- if (grepl("^\\d{4}_[01]\\d$", folder_date)) to_month_yyyy(folder_date) else folder_date
  
  # Pick most recent file matching regex + extension
  pick_most_recent <- function(dir, rx, ext) {
    files <- list.files(dir, pattern = paste0(rx, ".*\\.", ext, "$"),
                        full.names = TRUE, ignore.case = TRUE)
    if (!length(files)) return(character(0))
    info <- file.info(files)
    files[order(info$mtime, decreasing = TRUE)][1]
  }
  
  # Extract single integer from a capture group
  extract_count <- function(txt, rx, err) {
    m <- regexec(rx, txt, perl = TRUE, ignore.case = TRUE)
    hit <- regmatches(txt, m)
    if (!length(hit) || length(hit[[1]]) < 2) stop(err)
    as.integer(gsub(",", "", hit[[1]][2]))
  }
  
  # Split the full memo into all segments for the given month label
  split_month_segments <- function(full_txt, month_label) {
    t <- normalize_text(full_txt)
    start_rx <- paste0("In the month of(?:[[:space:]]+month of)?[[:space:]]+", month_label)
    starts <- gregexpr(start_rx, t, perl = TRUE, ignore.case = TRUE)[[1]]
    if (length(starts) == 1 && starts[1] == -1) return(character(0))
    next_markers <- gregexpr("In the month of", t, perl = TRUE, ignore.case = TRUE)[[1]]
    segs <- character(length(starts))
    for (i in seq_along(starts)) {
      s <- starts[i]
      next_after <- next_markers[next_markers > s]
      e <- if (length(next_after)) next_after[1] - 1L else nchar(t)
      segs[i] <- substr(t, s, e)
    }
    segs
  }
  
  # Prefer a segment that actually includes IDs + converted/reopened
  choose_target_segment <- function(segments) {
    if (!length(segments)) return("")
    id_hits  <- sapply(segments, function(s) length(regmatches(s, gregexpr("\\b(?:ID\\s*)?[0-9]{7,9}\\b", s, perl = TRUE))[[1]]))
    key_hits <- sapply(segments, function(s) grepl("\\b(converted|reopen(?:ed)?)\\b", s, ignore.case = TRUE))
    candidates <- which(id_hits > 0 & key_hits)
    if (length(candidates)) return(segments[candidates[which.max(id_hits[candidates])]])
    if (length(segments) >= 2) return(segments[2])
    segments[1]
  }
  
  # Grab 7–9 digit IDs with or without "ID" prefix
  find_ids <- function(text) {
    hits <- regmatches(text, gregexpr("\\b(?:ID\\s*)?[0-9]{7,9}\\b", text, perl = TRUE))[[1]]
    if (!length(hits)) character(0) else unique(gsub("\\D", "", hits))
  }
  
  # Sentence-level extraction with smart disambiguation when both keywords are present
  build_converted_reopen_df <- function(seg) {
    if (!nzchar(seg)) {
      return(data.frame(report_month = as.Date(character()), type = character(), resource_id = character(),
                        stringsAsFactors = FALSE))
    }
    
    # crude sentence split on periods; keeps last fragment if no trailing period
    sentences <- unlist(strsplit(seg, "\\.\\s+|\\.$"))
    sentences <- trimws(sentences)
    sentences <- sentences[nchar(sentences) > 0]
    
    ids_conv <- character(0)
    ids_rep  <- character(0)
    
    for (s in sentences) {
      has_conv <- grepl("\\bconverted\\b", s, ignore.case = TRUE)
      has_rep  <- grepl("\\breopen(?:ed)?\\b", s, ignore.case = TRUE)
      if (!has_conv && !has_rep) next
      
      if (has_conv && has_rep) {
        # Disambiguate by order: split at the first occurrence of the second keyword
        pos_conv <- regexpr("\\bconverted\\b", s, ignore.case = TRUE)[1]
        pos_rep  <- regexpr("\\breopen(?:ed)?\\b", s, ignore.case = TRUE)[1]
        
        if (pos_rep < pos_conv) {
          left  <- substr(s, 1, pos_conv - 1L)   # reopened side (before "converted")
          right <- substr(s, pos_conv, nchar(s)) # converted side
          ids_rep  <- c(ids_rep,  find_ids(left))
          ids_conv <- c(ids_conv, find_ids(right))
        } else {
          left  <- substr(s, 1, pos_rep - 1L)    # converted side (before "reopened")
          right <- substr(s, pos_rep, nchar(s))  # reopened side
          ids_conv <- c(ids_conv, find_ids(left))
          ids_rep  <- c(ids_rep,  find_ids(right))
        }
      } else if (has_conv) {
        ids_conv <- c(ids_conv, find_ids(s))
      } else if (has_rep) {
        ids_rep <- c(ids_rep, find_ids(s))
      }
    }
    
    ids_conv <- unique(ids_conv)
    ids_rep  <- unique(ids_rep)
    
    ids  <- c(ids_conv, ids_rep)
    type <- c(rep("converted", length(ids_conv)), rep("reopened", length(ids_rep)))
    
    if (!length(ids)) {
      return(data.frame(report_month = as.Date(character()), type = character(), resource_id = character(),
                        stringsAsFactors = FALSE))
    }
    
    data.frame(
      report_month = rep(report_month, length(ids)),
      type         = type,
      resource_id  = ids,
      stringsAsFactors = FALSE
    )
  }
  
  # ---------- Try PDF first ----------
  pdf_path <- pick_most_recent(subdir, pattern_pdf, "pdf")
  if (!length(pdf_path)) pdf_path <- pick_most_recent(subdir, "§3\\.1,[[:space:]]*3\\.3", "pdf")
  
  if (length(pdf_path)) {
    # Page 1 for headline counts
    onepage <- tempfile(fileext = ".pdf"); on.exit(unlink(onepage), add = TRUE)
    pdftools::pdf_subset(pdf_path, pages = 1, output = onepage)
    txt_p1 <- pdftools::pdf_text(onepage)[1]
    if (!nzchar(gsub("[[:space:]]+", "", txt_p1))) txt_p1 <- pdftools::pdf_ocr_text(onepage, pages = 1)[1]
    txt_p1 <- normalize_text(txt_p1)
    
    licensed_count_mdcps <- extract_count(
      txt_p1, "licensed ([0-9][0-9,]*) new homes\\b",
      paste0("Could not find licensed count on page 1 of: ", basename(pdf_path))
    )
    overdue_count_mdcps <- extract_count(
      txt_p1, "there[[:space:]]+were[[:space:]]+([0-9][0-9,]*)[[:space:]]+overdue\\b",
      paste0("Could not find overdue count on page 1 of: ", basename(pdf_path))
    )
    
    # Full memo for the month segment with IDs
    txt_full <- normalize_text(paste(pdftools::pdf_text(pdf_path), collapse = " "))
    segs <- split_month_segments(txt_full, month_label)
    seg  <- choose_target_segment(segs)
    mdcps_converted_reopen <- build_converted_reopen_df(seg)
    
    return(list(
      licensed_count_mdcps   = licensed_count_mdcps,
      overdue_count_mdcps    = overdue_count_mdcps,
      mdcps_converted_reopen = mdcps_converted_reopen,
      pdf_path = pdf_path,
      subdir   = subdir
    ))
  }
  
  # ---------- If no PDF, try DOCX ----------
  docx_path <- pick_most_recent(subdir, pattern_docx, "docx")
  if (!length(docx_path)) docx_path <- pick_most_recent(subdir, "§3\\.1,[[:space:]]*3\\.3", "docx")
  if (!length(docx_path)) {
    stop("No matching PDF ('", pattern_pdf, "') or DOCX ('", pattern_docx, "' or '§3.1, 3.3') found in: ", subdir)
  }
  
  doc <- officer::read_docx(docx_path)
  df  <- officer::docx_summary(doc)
  txt_full <- normalize_text(paste(na.omit(df$text), collapse = " "))
  
  licensed_count_mdcps <- extract_count(
    txt_full, "licensed ([0-9][0-9,]*) new homes\\b",
    paste0("Could not find licensed count in DOCX: ", basename(docx_path))
  )
  overdue_count_mdcps <- extract_count(
    txt_full, "there[[:space:]]+were[[:space:]]+([0-9][0-9,]*)[[:space:]]+overdue\\b",
    paste0("Could not find overdue count in DOCX: ", basename(docx_path))
  )
  
  segs <- split_month_segments(txt_full, month_label)
  seg  <- choose_target_segment(segs)
  mdcps_converted_reopen <- build_converted_reopen_df(seg)
  
  list(
    licensed_count_mdcps   = licensed_count_mdcps,
    overdue_count_mdcps    = overdue_count_mdcps,
    mdcps_converted_reopen = mdcps_converted_reopen,
    docx_path = docx_path,
    subdir    = subdir
  )
}

########################################
# 3.1 non-relative - create file for PC review
########################################

# Create list of converted and reopened resource ids for PC to review.
# Save to monthly processed folders, and create a cumulative file for sending
# to PC

# One-shot runner: monthly exports + cumulative export
make_converted_reopened_exports <- function(
    months,
    base_path_pdf,
    folder_data,
    commitment              = "3.1",
    commitment_description  = "Non Relatives Homes Licensed",
    output_description      = "Non Relatives Homes Converted or Reopened",  # <-- new
    setup_folders_fun       = NULL,
    parser_fun              = get_mdcps_3_1_counts
) {
  # ----- helpers -----
  add_pc_cols <- function(df) {
    df %>% mutate(pc_approved_yes = "", pc_notes = "")
  }
  safe_filename <- function(x) {
    # replace characters Windows/macOS don't like in filenames
    gsub('[\\/:*?"<>|]+', "_", x)
  }
  
  default_setup <- function(m) {
    month_dir <- file.path(folder_data, m)
    processed <- file.path(month_dir, "processed")
    raw       <- file.path(month_dir, "raw")
    dir.create(processed, recursive = TRUE, showWarnings = FALSE)
    dir.create(raw,       recursive = TRUE, showWarnings = FALSE)
    list(folder_date = m, folder_processed = processed, folder_raw = raw)
  }
  if (is.null(setup_folders_fun)) setup_folders_fun <- default_setup
  
  schema <- tibble(
    report_month    = as.Date(character()),
    type            = character(),
    resource_id     = character(),
    pc_approved_yes = character(),
    pc_notes        = character()
  )
  
  monthly_files <- character(0)
  cum_list <- list()
  run_date <- Sys.Date()   # 2025-09-21 in your example
  
  for (m in months) {
    setup <- setup_folders_fun(m)
    folder_date      <- setup$folder_date      # e.g., "2024_03"
    folder_processed <- setup$folder_processed
    
    # parse THIS month explicitly
    res <- parser_fun(base_path_pdf, folder_date = m)
    
    # build the frame we actually write
    cr <- res$mdcps_converted_reopen
    if (is.null(cr)) cr <- schema[0, c("report_month", "type", "resource_id")]
    out_df <- cr %>%
      mutate(
        report_month = as.Date(report_month),
        resource_id  = as.character(resource_id)
      ) %>%
      add_pc_cols()
    
    if (nrow(out_df)) {
      ok <- all(format(out_df$report_month, "%Y_%m") == m)
      if (!ok) warning("report_month mismatch in ", m, " — writing anyway.")
    }
    
    # monthly path: .../data/YYYY_MM/processed/YYYY-MM-DD/<file>.xlsx
    run_folder <- file.path(folder_processed, format(run_date, "%Y-%m-%d"))
    dir.create(run_folder, recursive = TRUE, showWarnings = FALSE)
    
    # ********** FILE NAME STEM (what you asked for) **********
    # "2024_03 - 3.1 - Non Relatives Homes Converted or Reopened - 2025-09-21.xlsx"
    file_stem <- paste(
      folder_date,
      commitment,
      output_description,
      format(run_date, "%Y-%m-%d"),
      sep = " - "
    )
    monthly_file <- file.path(run_folder, paste0(safe_filename(file_stem), ".xlsx"))
    # *******************************************************
    
    # write monthly file
    wb <- createWorkbook()
    addWorksheet(wb, "converted_reopened")
    writeData(wb, "converted_reopened", out_df)
    
    rid_col <- which(names(out_df) == "resource_id")
    rm_col  <- which(names(out_df) == "report_month")
    if (nrow(out_df) > 0 && length(rid_col) == 1) {
      addStyle(wb, "converted_reopened", createStyle(numFmt = "@"),
               rows = 2:(nrow(out_df) + 1), cols = rid_col, gridExpand = TRUE, stack = TRUE)
    }
    if (nrow(out_df) > 0 && length(rm_col) == 1) {
      addStyle(wb, "converted_reopened", createStyle(numFmt = "yyyy-mm-dd"),
               rows = 2:(nrow(out_df) + 1), cols = rm_col, gridExpand = TRUE, stack = TRUE)
    }
    saveWorkbook(wb, monthly_file, overwrite = TRUE)
    message("Wrote ", basename(monthly_file), " (rows: ", nrow(out_df), ")")
    
    monthly_files <- c(monthly_files, monthly_file)
    cum_list[[length(cum_list) + 1L]] <- out_df
  }
  
  # cumulative to MAIN data folder
  cum_df <- if (length(cum_list)) {
    dplyr::bind_rows(c(list(schema), purrr::compact(cum_list)))
  } else {
    schema
  }
  
  # Optional: give the cumulative a parallel, human-readable name too
  cum_stem <- paste(
    folder_date,
    commitment,
    output_description,
    format(run_date, "%Y-%m-%d"),
    sep = " - "
  )
  cum_file <- file.path(folder_data, paste0(safe_filename(cum_stem), ".xlsx"))
  
  wb <- createWorkbook()
  addWorksheet(wb, "converted_reopened")
  writeData(wb, "converted_reopened", cum_df)
  
  rid_col <- which(names(cum_df) == "resource_id")
  rm_col  <- which(names(cum_df) == "report_month")
  if (nrow(cum_df) > 0 && length(rid_col) == 1) {
    addStyle(wb, "converted_reopened", createStyle(numFmt = "@"),
             rows = 2:(nrow(cum_df) + 1), cols = rid_col, gridExpand = TRUE, stack = TRUE)
  }
  if (nrow(cum_df) > 0 && length(rm_col) == 1) {
    addStyle(wb, "converted_reopened", createStyle(numFmt = "yyyy-mm-dd"),
             rows = 2:(nrow(cum_df) + 1), cols = rm_col, gridExpand = TRUE, stack = TRUE)
  }
  saveWorkbook(wb, cum_file, overwrite = TRUE)
  message("Cumulative written: ", basename(cum_file))
  
  invisible(list(
    cumulative_path = cum_file,
    cumulative_df   = cum_df,
    monthly_files   = monthly_files
  ))
}

########################################
# Copy current and prior files from sharefile to my folder_raw folder ----
########################################

# The function will find files in targeted folder(s) whose timeframes (based
# on file name pattern, like "...2024 Q2 ...") match the timeframe specified in 
# folder_date and occurred prior to it but within the same year (e.g., "2024 Q2"
# and "2024 Q1"). It then copies these files into the current folder_raw to
# so the files from multiple quarters can be merged to create a cumulative file. 
# It works if the specified path has no subfolders, has quarterly or monthly subfolders
# (like 2024_Q2 or 2024 01), or has quarterly or monthly subfolders with 
# date-named subfolders (like 2024_Q2/2024-03-15).

copy_sharefile_files <- function(source_path) {
  # --- Step 1: Look for a subfolder in source_path matching folder_date in various formats ---
  
  # List all immediate subdirectories of source_path
  subfolders <- list.dirs(source_path, full.names = TRUE, recursive = FALSE)
  
  # Determine which format folder_date is: quarterly ("YYYY_QN") or monthly ("YYYY_MM")
  if (grepl("_Q", folder_date, ignore.case = TRUE)) {
    # Quarterly folder_date: e.g., "2024_Q1"
    year_current   <- sub("^(\\d{4})_Q[1-4]$", "\\1", folder_date, ignore.case = TRUE)
    quarter_current <- sub("^\\d{4}_Q([1-4])$", "\\1", folder_date, ignore.case = TRUE)
    # Build a regex that will match any of:
    # "YYYY_QN", "YYYY QN", "Q1_YYYY", or "Q1 YYYY"
    pattern <- paste0("(?i)^(?:", year_current, "[_\\s]+Q", quarter_current, "|Q", quarter_current, "[_\\s]+", year_current, ")$")
    
  } else if (grepl("^\\d{4}_[0-9]{2}$", folder_date)) {
    # Monthly folder_date: e.g., "2024_01"
    year_current <- sub("^(\\d{4})_[0-9]{2}$", "\\1", folder_date)
    month_current <- sub("^\\d{4}_([0-9]{2})$", "\\1", folder_date)
    # Build a regex that matches either "YYYY_MM" or "YYYY-MM"
    pattern <- paste0("(?i)^", year_current, "[-_]", month_current, "$")
    
  } else {
    stop("folder_date format not recognized. It should be either 'YYYY_QN' or 'YYYY_MM'.")
  }
  
  # Look for a candidate subfolder whose basename matches the pattern.
  candidate <- subfolders[grepl(pattern, basename(subfolders), perl = TRUE)]
  if (length(candidate) > 0 && dir.exists(candidate[1])) {
    target_folder <- candidate[1]
    message("Found subfolder matching folder_date: ", target_folder)
  } else {
    message("No subfolder matching folder_date found in ", source_path, 
            ". Files will be processed directly in the source folder.")
    target_folder <- source_path
  }
  
  # --- Step 2: Within target_folder, check for date-named subdirectories (YYYY-MM-DD) ---
  subfolders_date <- list.dirs(target_folder, full.names = TRUE, recursive = FALSE)
  if (length(subfolders_date) > 0) {
    subfolder_names <- basename(subfolders_date)
    subfolder_dates <- as.Date(subfolder_names, format = "%Y-%m-%d")
    
    valid_idx <- which(!is.na(subfolder_dates))
    if (length(valid_idx) > 0) {
      # Select the subfolder with the most recent date.
      latest_idx <- valid_idx[which.max(subfolder_dates[valid_idx])]
      target_folder <- subfolders_date[latest_idx]
      message("Found date-named subfolders. Processing files from the most recent folder: ", target_folder)
    } else {
      message("No valid date-named subdirectories found in ", target_folder, ".")
    }
  } else {
    message("No subdirectories found in ", target_folder, ".")
  }
  
  # --- Step 3: Ensure the destination folder (folder_raw) exists ---
  if (!dir.exists(folder_raw)) {
    dir.create(folder_raw, recursive = TRUE)
    message("Created destination folder: ", folder_raw)
  }
  
  # --- Step 4: List all .xlsx files in the target folder ---
  file_list <- list.files(target_folder, pattern = "\\.xlsx$", 
                          ignore.case = TRUE, full.names = TRUE)
  
  # Prepare vectors to store matching files and processing messages.
  files_to_copy <- character(0)
  message_lines <- character(0)
  
  # --- Step 5: Process files based on folder_date format (quarterly or monthly) ---
  if (grepl("_Q", folder_date, ignore.case = TRUE)) {
    ## Quarterly processing: folder_date is expected as "YYYY_QN" (e.g., "2024_Q2")
    year_current    <- as.integer(sub("^(\\d{4})_Q[1-4]$", "\\1", folder_date, ignore.case = TRUE))
    quarter_current <- as.integer(sub("^\\d{4}_Q([1-4])$", "\\1", folder_date, ignore.case = TRUE))
    
    # Regex to match either "YYYY_QN" (or "YYYY QN") or "Qn_YYYY" (or "Qn YYYY")
    pattern <- "(?i)(?:(\\d{4})[_\\s]*Q([1-4])|Q([1-4])[_\\s]*(\\d{4}))"
    
    for (file in file_list) {
      file_base <- basename(file)
      m_exec <- regexec(pattern, file_base, perl = TRUE)
      capture <- regmatches(file_base, m_exec)[[1]]
      
      if (length(capture) > 0 && capture[1] != "") {
        # For Alternative 1, capture[2] is the year and capture[3] is the quarter.
        # For Alternative 2, capture[4] is the quarter and capture[5] is the year.
        if (nzchar(capture[2])) {
          file_year    <- as.integer(capture[2])
          file_quarter <- as.integer(capture[3])
        } else {
          file_quarter <- as.integer(capture[4])
          file_year    <- as.integer(capture[5])
        }
        
        if (file_year == year_current && file_quarter <= quarter_current) {
          files_to_copy <- c(files_to_copy, file)
          message_lines <- c(message_lines, 
                             paste("- MATCH (Quarter", file_quarter, file_year, ") -", file_base))
        } else {
          message_lines <- c(message_lines, 
                             paste("- NOT A MATCH (found Quarter", file_quarter, file_year, ") -", file_base))
        }
      } else {
        message_lines <- c(message_lines, 
                           paste("- NO TIMEFRAME PATTERN FOUND -", file_base))
      }
    }
    
  } else if (grepl("^\\d{4}_[0-9]{2}$", folder_date)) {
    ## Monthly processing: folder_date is expected as "YYYY_MM" (e.g., "2024_03")
    year_current  <- as.integer(sub("^(\\d{4})_[0-9]{2}$", "\\1", folder_date))
    month_current <- as.integer(sub("^\\d{4}_([0-9]{2})$", "\\1", folder_date))
    
    # Look for a pattern like "YYYY_MM" or "YYYY-MM" within file names.
    pattern <- "([0-9]{4})[-_](0[1-9]|1[0-2])"
    
    for (file in file_list) {
      file_base <- basename(file)
      m_exec <- regexec(pattern, file_base, perl = TRUE)
      capture <- regmatches(file_base, m_exec)[[1]]
      
      if (length(capture) > 0) {
        file_year  <- as.integer(capture[2])
        file_month <- as.integer(capture[3])
        
        if (file_year == year_current && file_month <= month_current) {
          files_to_copy <- c(files_to_copy, file)
          message_lines <- c(message_lines, 
                             paste("- MATCH (Month", file_month, file_year, ") -", file_base))
        } else {
          message_lines <- c(message_lines, 
                             paste("- NOT A MATCH (found Month", file_month, file_year, ") -", file_base))
        }
      } else {
        message_lines <- c(message_lines, 
                           paste("- NO TIMEFRAME PATTERN FOUND -", file_base))
      }
    }
    
  } else {
    stop("folder_date format not recognized. It should be either 'YYYY_QN' or 'YYYY_MM'.")
  }
  
  # --- Step 6: Check if matching files already exist in folder_raw and copy only new files ---
  if (length(files_to_copy) > 0) {
    dest_paths <- file.path(folder_raw, basename(files_to_copy))
    new_files <- files_to_copy[!file.exists(dest_paths)]
    
    if (length(new_files) == 0) {
      message("No new files to copy. All files which matched already exist in ", folder_raw, ".")
    } else {
      success <- file.copy(from = new_files, to = folder_raw, overwrite = TRUE)
      if (!all(success)) {
        warning("Some files failed to copy.")
      }
    }
  } else {
    message("No files to copy based on the timeframe criteria.")
  }
  
  # --- Step 7: Output the processing summary ---
  message("Files processed:")
  for (line in message_lines) {
    message(line)
  }
}

########################################
# Compare two data frames and save multi-sheet comparison workbook ----
########################################

# Purpose:
# Compare two data frames (typically yours vs. Eric's) on specified ID columns,
# optionally filter each data frame before comparison, and create multiple
# comparison sheets based on different metric filters. Saves results as a
# multi-sheet Excel workbook.
#
# Args:
#   df1          - Your data frame (typically the main/project data)
#   df2          - Comparison data frame (typically Eric's data)
#   id_cols      - Character vector of column names to join on
#                  (e.g., c("child_id", "custody_sequence_number", "placement_id"))
#   df1_filter   - Optional filter expression for df1 (use quo() to quote)
#                  Can include multiple conditions: quo(quarter == "2025 Q1" & month == "January")
#   df2_filter   - Optional filter expression for df2 (use quo() to quote)
#                  Can include multiple conditions: quo(quarter == "2025 Q1" & month == "January")
#   metrics      - Named list of filter expressions for creating multiple comparison sheets
#                  Each element creates one sheet. Use quo() for each expression.
#                  Example: list("THV Received" = quo(thv_received == TRUE))
#   extra_fields - Character vector of additional field names from df1 to include
#                  in the comparison output (e.g., c("custody_start", "placement_start"))
#   df1_label    - Label for df1 in output columns (default: "mine")
#   df2_label    - Label for df2 in output columns (default: "eric")
#   save_file    - Logical; if TRUE, saves workbook using save_workbook_to_folder_run()
#
# Returns:
#   Named list of comparison data frames (one per metric), invisibly
#
# Usage:
#   # Simple comparison with one metric
#   compare_and_save(
#     df1 = thv_children,
#     df2 = eric_detail_df,
#     id_cols = c("child_id", "custody_sequence_number"),
#     df2_filter = quo(quarter == folder_date_quarter),
#     metrics = list("THV Received" = quo(thv_received == TRUE)),
#     df1_label = "kurt",
#     df2_label = "eric"
#   )
#
#   # Multiple metrics with extra fields
#   compare_and_save(
#     df1 = thv_children,
#     df2 = eric_detail_df,
#     id_cols = c("reporting_month", "child_id", "custody_sequence_number", "placement_id"),
#     df2_filter = quo(quarter == folder_date_quarter & reporting_month == "January"),
#     metrics = list(
#       "THV Received" = quo(thv_received == TRUE),
#       "THV Eligible" = quo(thv_eligible == TRUE)
#     ),
#     extra_fields = c("custody_start", "custody_end", "placement_start", "placement_end"),
#     df1_label = "kurt",
#     df2_label = "eric"
#   )

compare_and_save <- function(
  df1,
  df2,
  id_cols,
  df1_filter = NULL,
  df2_filter = NULL,
  metrics = NULL,
  extra_fields = NULL,
  df1_label = "mine",
  df2_label = "eric",
  save_file = TRUE
) {

  # ---- validate inputs ----
  if (!is.data.frame(df1) || !is.data.frame(df2)) {
    stop("df1 and df2 must be data frames.")
  }

  if (!all(id_cols %in% names(df1))) {
    stop("Not all id_cols found in df1: ", paste(setdiff(id_cols, names(df1)), collapse = ", "))
  }

  if (!all(id_cols %in% names(df2))) {
    stop("Not all id_cols found in df2: ", paste(setdiff(id_cols, names(df2)), collapse = ", "))
  }

  if (!is.null(extra_fields) && !all(extra_fields %in% names(df1))) {
    stop("Not all extra_fields found in df1: ", paste(setdiff(extra_fields, names(df1)), collapse = ", "))
  }

  if (is.null(metrics)) {
    stop("metrics must be provided as a named list of filter expressions.")
  }

  # ---- apply pre-filters ----
  df1_filtered <- df1
  df2_filtered <- df2

  if (!is.null(df1_filter)) {
    df1_filtered <- df1 %>% filter(!!df1_filter)
    message("Applied df1_filter: ", nrow(df1_filtered), " rows remaining (from ", nrow(df1), ")")
  }

  if (!is.null(df2_filter)) {
    df2_filtered <- df2 %>% filter(!!df2_filter)
    message("Applied df2_filter: ", nrow(df2_filtered), " rows remaining (from ", nrow(df2), ")")
  }

  # ---- build comparison for each metric ----
  comparison_sheets <- list()

  for (metric_name in names(metrics)) {
    metric_filter <- metrics[[metric_name]]

    # Filter df1 by metric
    df1_metric <- df1_filtered %>%
      filter(!!metric_filter) %>%
      select(all_of(c(id_cols, extra_fields))) %>%
      distinct() %>%
      mutate(!!paste0("in_", df1_label) := TRUE)

    # Filter df2 by metric
    df2_metric <- df2_filtered %>%
      filter(!!metric_filter) %>%
      select(all_of(id_cols)) %>%
      distinct() %>%
      mutate(!!paste0("in_", df2_label) := TRUE)

    # Full join to compare
    comparison <- full_join(
      df1_metric,
      df2_metric,
      by = id_cols
    ) %>%
      mutate(
        !!paste0("in_", df1_label) := coalesce(!!sym(paste0("in_", df1_label)), FALSE),
        !!paste0("in_", df2_label) := coalesce(!!sym(paste0("in_", df2_label)), FALSE),
        in_both = !!sym(paste0("in_", df1_label)) & !!sym(paste0("in_", df2_label))
      ) %>%
      arrange(across(all_of(id_cols)))

    comparison_sheets[[metric_name]] <- comparison

    # Report counts
    n_df1 <- sum(comparison[[paste0("in_", df1_label)]])
    n_df2 <- sum(comparison[[paste0("in_", df2_label)]])
    n_both <- sum(comparison$in_both)

    message(sprintf(
      "%s: %s=%d, %s=%d, both=%d",
      metric_name, df1_label, n_df1, df2_label, n_df2, n_both
    ))
  }

  # ---- save workbook ----
  if (save_file) {
    # Requires folder_run to be set up
    if (!exists("folder_run", envir = .GlobalEnv)) {
      if (exists("make_data_run_folder", mode = "function")) {
        make_data_run_folder()
      } else {
        stop("folder_run not found. Call make_data_run_folder() first (or create folder_run).")
      }
    }

    # Get global variables for filename
    folder_run <- get("folder_run", envir = .GlobalEnv)
    folder_date <- if (exists("folder_date", envir = .GlobalEnv)) {
      get("folder_date", envir = .GlobalEnv)
    } else {
      stop("folder_date not found in global environment.")
    }
    commitment <- if (exists("commitment", envir = .GlobalEnv)) {
      get("commitment", envir = .GlobalEnv)
    } else {
      stop("commitment not found in global environment.")
    }

    # Build filename: folder_date_commitment_compare_YYYY-MM-DD.xlsx
    out_file <- file.path(
      folder_run,
      paste0(folder_date, "_", commitment, "_compare_", Sys.Date(), ".xlsx")
    )

    # Create and save workbook
    if (!requireNamespace("openxlsx", quietly = TRUE)) {
      stop("The openxlsx package is required. Install it with: install.packages('openxlsx')")
    }

    wb <- openxlsx::createWorkbook()
    hdr_style <- openxlsx::createStyle(textDecoration = "bold", halign = "left", valign = "top")

    for (sheet_name in names(comparison_sheets)) {
      df <- comparison_sheets[[sheet_name]]

      # Sanitize sheet name (Excel has 31 char limit and some forbidden characters)
      clean_sheet_name <- substr(gsub('[/\\\\:*?\\[\\]]+', "-", sheet_name), 1, 31)

      openxlsx::addWorksheet(wb, clean_sheet_name)
      openxlsx::writeData(wb, clean_sheet_name, df, headerStyle = hdr_style, borders = "none")
      openxlsx::setColWidths(wb, clean_sheet_name, cols = 1:ncol(df), widths = "auto")
    }

    openxlsx::saveWorkbook(wb, out_file, overwrite = TRUE)
    message("Saved comparison workbook: ", out_file)
  }

  invisible(comparison_sheets)
}

########################################
# Run analysis for all quarters YTD and save cumulative workbook ----
########################################

# Purpose:
# Runs an analysis function for each quarter from Q1 to the current quarter within
# the same year, collecting results for each quarter and binding them together to
# create a year-to-date (YTD) cumulative workbook. This ensures consistent analysis
# logic across all quarters and avoids schema/calculation drift.
#
# Args:
#   folder_date  - Current quarter in format "YYYY_QX" (e.g., "2025_Q2")
#                  Must exist in global environment if not provided
#   analysis_fn  - Function that performs analysis for ONE quarter and returns
#                  a named list of data frames (one per sheet)
#   data_root    - Root directory for data (default: "data")
#
# The analysis_fn should:
#   - Load raw data files (which don't change by quarter)
#   - Filter/process data based on the current iteration's quarter
#   - Use global variables like folder_date_quarter that are auto-set per iteration
#   - Return a named list where names = sheet names, values = data frames
#
# Returns:
#   Named list of YTD data frames (invisibly), one per sheet, containing
#   data from Q1 through current quarter
#
# Side effects:
#   - Temporarily modifies global variables during iteration (restored after)
#   - Saves individual quarter files for all quarters EXCEPT the final quarter
#   - Saves YTD workbook for the final quarter using save_workbook_to_folder_run()
#
# File saving behavior:
#   - Q1 run: Saves Q1 file to data/2025_Q1/processed/YYYY-MM-DD/
#   - Q2 run: Saves Q1 individual + Q2 YTD (containing Q1+Q2) to data/2025_Q2/processed/YYYY-MM-DD/
#   - Q3 run: Saves Q1 individual + Q2 individual + Q3 YTD (Q1+Q2+Q3) to data/2025_Q3/processed/YYYY-MM-DD/
#
# Usage:
#   my_setup <- setup_folders("2025_Q2")
#
#   sheets_ytd <- run_ytd_analysis(
#     folder_date = folder_date,
#     analysis_fn = function() {
#       # Load raw data
#       custody_df <- load_sharefile_file(ERIC_BASE, "custody", "custody_df")
#       placements_df <- load_sharefile_file(ERIC_BASE, "placements", "placements_df")
#
#       # Process (uses auto-set folder_date_quarter for this iteration)
#       custody_joined_plans_aug <- custody_joined_plans %>%
#         filter(reporting_qtr == folder_date_quarter)
#
#       thv_children <- custody_joined_plans_aug %>%
#         filter(thv_eligible, days_placed >= 30)
#
#       # Return sheets for this quarter
#       list(
#         "Population Details" = thv_children,
#         "Quarter Summary" = create_thv_summary(thv_children, "quarter"),
#         "Monthly Summary" = create_thv_summary(thv_children, c("quarter", "report_month"))
#       )
#     }
#   )

run_ytd_analysis <- function(
  folder_date = NULL,
  analysis_fn,
  data_root = "data"
) {

  # ---- validate inputs ----
  if (is.null(folder_date)) {
    if (exists("folder_date", envir = .GlobalEnv)) {
      folder_date <- get("folder_date", envir = .GlobalEnv)
    } else {
      stop("folder_date must be provided or exist in global environment.")
    }
  }

  if (!is.function(analysis_fn)) {
    stop("analysis_fn must be a function that returns a named list of data frames.")
  }

  # ---- parse folder_date ----
  if (!grepl("^\\d{4}_Q[1-4]$", folder_date)) {
    stop("folder_date must be in format 'YYYY_QX' (e.g., '2025_Q2').")
  }

  year <- substr(folder_date, 1, 4)
  current_q <- as.integer(substr(folder_date, 7, 7))

  quarters_to_process <- paste0(year, "_Q", 1:current_q)

  message("Running YTD analysis for: ", paste(quarters_to_process, collapse = ", "))

  # ---- save original global state ----
  globals_to_restore <- list()
  global_vars <- c("folder_date", "reporting_period_start", "reporting_period_end",
                   "folder_date_readable", "folder_date_quarter", "cy_start",
                   "folder_raw", "folder_processed", "folder_output")

  for (var in global_vars) {
    if (exists(var, envir = .GlobalEnv)) {
      globals_to_restore[[var]] <- get(var, envir = .GlobalEnv)
    }
  }

  # Ensure restoration happens even if there's an error
  on.exit({
    for (var in names(globals_to_restore)) {
      assign(var, globals_to_restore[[var]], envir = .GlobalEnv)
    }
    message("Restored original global variables")
  }, add = TRUE)

  # ---- process each quarter ----
  all_quarters_results <- list()

  for (qtr in quarters_to_process) {
    message("\n--- Processing quarter: ", qtr, " ---")

    # Set up folders and globals for this quarter
    # This updates folder_date, folder_date_quarter, folder_raw, folder_processed, etc.
    setup_result <- setup_folders(qtr, assign_globals = TRUE, data_root = data_root)

    # Verify required file paths exist
    if (!exists("folder_raw", envir = .GlobalEnv)) {
      stop("folder_raw was not set by setup_folders for quarter: ", qtr)
    }

    folder_raw_qtr <- get("folder_raw", envir = .GlobalEnv)
    folder_processed_qtr <- get("folder_processed", envir = .GlobalEnv)

    # Check that raw data folder exists
    if (!dir.exists(folder_raw_qtr)) {
      stop("Raw data folder does not exist for quarter ", qtr, ": ", folder_raw_qtr,
           "\nEnsure data has been loaded for all quarters from Q1 to current quarter.")
    }

    # Run the user's analysis function for this quarter
    tryCatch({
      quarter_results <- analysis_fn()

      # Validate results
      if (!is.list(quarter_results) || is.null(names(quarter_results))) {
        stop("analysis_fn must return a named list of data frames for quarter: ", qtr)
      }

      if (!all(sapply(quarter_results, is.data.frame))) {
        stop("All elements returned by analysis_fn must be data frames for quarter: ", qtr)
      }

      # Store results for this quarter
      all_quarters_results[[qtr]] <- quarter_results

      message("✓ Successfully processed quarter: ", qtr)
      for (sheet_name in names(quarter_results)) {
        message("  - ", sheet_name, ": ", nrow(quarter_results[[sheet_name]]), " rows")
      }

      # ---- Save individual quarter file (only for non-final quarters) ----
      # For the current/final quarter, we'll save the YTD file instead
      is_final_quarter <- (qtr == quarters_to_process[length(quarters_to_process)])

      if (!is_final_quarter) {
        # Clear any stale folder_run, then create run folder for THIS quarter
        if (exists("folder_run", envir = .GlobalEnv)) {
          rm("folder_run", envir = .GlobalEnv)
        }

        if (exists("make_data_run_folder", mode = "function")) {
          make_data_run_folder()
        } else {
          stop("Cannot create folder_run. make_data_run_folder() function not found.")
        }

        # Save THIS quarter's file to its own folder
        # e.g., data/2025_Q1/processed/2025-10-15/2025_Q1_5.1.b_...xlsx
        save_workbook_to_folder_run(quarter_results)
        message("✓ Saved ", qtr, " individual file to: ", get("folder_run", envir = .GlobalEnv))
      } else {
        message("  (Will save YTD file for ", qtr, " after binding all quarters)")
      }

    }, error = function(e) {
      stop("Error processing quarter ", qtr, ": ", e$message)
    })
  }

  # ---- bind all quarters by sheet name ----
  message("\n--- Binding all quarters together ---")

  # Get all unique sheet names across all quarters
  all_sheet_names <- unique(unlist(lapply(all_quarters_results, names)))

  ytd_sheets <- list()
  for (sheet_name in all_sheet_names) {
    # Collect this sheet from each quarter
    sheet_data_list <- lapply(all_quarters_results, function(qtr_result) {
      if (sheet_name %in% names(qtr_result)) {
        qtr_result[[sheet_name]]
      } else {
        NULL
      }
    })

    # Remove NULLs (sheet didn't exist in some quarter)
    sheet_data_list <- Filter(Negate(is.null), sheet_data_list)

    if (length(sheet_data_list) == 0) {
      warning("Sheet '", sheet_name, "' has no data in any quarter. Skipping.")
      next
    }

    # Bind rows across quarters
    ytd_sheets[[sheet_name]] <- bind_rows(sheet_data_list)

    message("  - ", sheet_name, ": ", nrow(ytd_sheets[[sheet_name]]), " total rows (YTD)")
  }

  # ---- restore original folder context and save ----
  # Restore original globals before saving (so file goes to correct location)
  for (var in names(globals_to_restore)) {
    assign(var, globals_to_restore[[var]], envir = .GlobalEnv)
  }

  message("\n--- Saving YTD workbook ---")

  # IMPORTANT: Clear any stale folder_run from previous runs
  # Then create fresh folder_run for the CURRENT quarter
  if (exists("folder_run", envir = .GlobalEnv)) {
    rm("folder_run", envir = .GlobalEnv)
  }

  # This creates: data/<current_quarter>/processed/<YYYY-MM-DD>/
  if (exists("make_data_run_folder", mode = "function")) {
    make_data_run_folder()
  } else {
    stop("Cannot create folder_run. make_data_run_folder() function not found.")
  }

  folder_run_final <- get("folder_run", envir = .GlobalEnv)
  message("Saving to: ", folder_run_final)

  save_workbook_to_folder_run(ytd_sheets)

  message("✓ YTD analysis complete!")

  invisible(ytd_sheets)
}

########################################
# Format quarterly monitoring table for R Markdown memos ----
########################################

#' Create formatted quarterly monitoring table
#'
#' @param data Data frame with columns: quarter, metric, count_monitor, percent_monitor
#' @param footnote Optional footnote text to display below table
#'
#' @return A flextable object formatted for Word output
#'
#' @examples
#' table_quarter_monitor_only(thv_children_quarter_long)
#' table_quarter_monitor_only(thv_children_quarter_long, footnote = "Note: Children are reported...")

table_quarter_monitor_only <- function(data, footnote = NULL, table_title = "Table 1. Quarterly Results.") {

  if (!requireNamespace("flextable", quietly = TRUE)) {
    stop("flextable package required. Install with: install.packages('flextable')")
  }

  # Prepare data: rename columns for display
  df <- data %>%
    select(quarter, metric, count_monitor, percent_monitor) %>%
    mutate(
      `#` = count_monitor,
      `%` = ifelse(is.na(percent_monitor), "", paste0(percent_monitor, "%"))
    ) %>%
    select(quarter, metric, `#`, `%`)

  # Create flextable
  ft <- flextable::flextable(df) %>%
    # Remove default column labels for quarter and metric in header row 2
    flextable::set_header_labels(quarter = "", metric = "") %>%
    # Add header row with "Monitor" spanning columns 3-4
    flextable::add_header_row(
      values = c("", "", "Monitor"),
      colwidths = c(1, 1, 2)
    ) %>%
    # Merge quarter cells vertically
    flextable::merge_v(j = 1, part = "body") %>%
    # Align columns
    flextable::align(align = "left", part = "all") %>%
    flextable::align(j = c(3, 4), align = "center", part = "body") %>%
    flextable::align(j = c(3, 4), align = "center", part = "header") %>%
    flextable::valign(j = 1, valign = "center", part = "body") %>%
    # Set fonts
    flextable::font(fontname = "Calibri", part = "all") %>%
    flextable::fontsize(size = 11, part = "all") %>%
    # Column widths
    flextable::width(j = 1, width = 0.8) %>%
    flextable::width(j = 2, width = 4.0) %>%
    flextable::width(j = 3, width = 0.6) %>%
    flextable::width(j = 4, width = 0.6) %>%
    # Borders - all width = 1
    flextable::border_remove() %>%
    flextable::hline_top(border = officer::fp_border(width = 1), part = "header") %>%
    flextable::hline_bottom(border = officer::fp_border(width = 1), part = "header") %>%
    flextable::hline_bottom(border = officer::fp_border(width = 1), part = "body")

  # Add zebra striping and quarter separators
  quarters <- unique(df$quarter)
  if (length(quarters) > 1) {
    # Multiple quarters: add separators between quarters
    for (i in 1:(length(quarters) - 1)) {
      last_row <- max(which(df$quarter == quarters[i]))
      ft <- ft %>%
        flextable::hline(i = last_row, border = officer::fp_border(width = 1))
    }
  }

  # Zebra striping (alternating rows within each quarter)
  for (q in quarters) {
    q_rows <- which(df$quarter == q)
    shade_rows <- q_rows[seq(2, length(q_rows), 2)]  # Every other row
    if (length(shade_rows) > 0) {
      ft <- ft %>%
        flextable::bg(i = shade_rows, bg = "#F0F0F0", part = "body")
    }
  }

  # Add table title with bold formatting in Calibri 11
  ft <- ft %>%
    flextable::set_caption(
      caption = as_paragraph(
        as_chunk(table_title,
                 props = officer::fp_text(font.family = "Calibri", font.size = 11, bold = TRUE))
      ),
      style = "Normal",
      align_with_table = FALSE,
      fp_p = officer::fp_par(text.align = "left", padding.bottom = 6)
    )

  # Add footnote if provided (not italicized)
  if (!is.null(footnote)) {
    ft <- ft %>%
      flextable::add_footer_lines(footnote) %>%
      flextable::align(align = "left", part = "footer") %>%
      flextable::font(fontname = "Calibri", part = "footer") %>%
      flextable::fontsize(size = 11, part = "footer")
  }

  ft
}

########################################
# Format quarterly + monthly monitoring table for R Markdown memos ----
########################################

#' Create formatted quarterly/monthly monitoring table
#'
#' @param data Data frame with columns: quarter, report_month, metric, count_monitor, percent_monitor
#' @param footnote Optional footnote text to display below table
#'
#' @return A flextable object formatted for Word output
#'
#' @examples
#' table_quarter_month_monitor_only(thv_children_month_long)

table_quarter_month_monitor_only <- function(data, footnote = NULL, table_title = "Table 1. Monthly Results.") {

  if (!requireNamespace("flextable", quietly = TRUE)) {
    stop("flextable package required. Install with: install.packages('flextable')")
  }

  # Prepare data: rename columns and format months
  df <- data %>%
    mutate(
      Month = format(as.Date(paste0("2025-", match(report_month, month.name), "-01")), "%b"),
      `#` = count_monitor,
      `%` = ifelse(is.na(percent_monitor), "", paste0(percent_monitor, "%"))
    ) %>%
    select(quarter, Month, metric, `#`, `%`)

  # Create flextable
  ft <- flextable::flextable(df) %>%
    # Remove default column labels for quarter, month, and metric in header row 2
    flextable::set_header_labels(quarter = "", Month = "", metric = "") %>%
    # Add header row with "Monitor" spanning columns 4-5
    flextable::add_header_row(
      values = c("", "", "", "Monitor"),
      colwidths = c(1, 1, 1, 2)
    ) %>%
    # Merge quarter and month cells vertically within each quarter/month group
    flextable::merge_v(j = 1, part = "body") %>%
    flextable::merge_v(j = 2, part = "body") %>%
    # Align columns
    flextable::align(align = "left", part = "all") %>%
    flextable::align(j = c(4, 5), align = "center", part = "body") %>%
    flextable::align(j = c(4, 5), align = "center", part = "header") %>%
    flextable::valign(j = c(1, 2), valign = "center", part = "body") %>%
    # Set fonts
    flextable::font(fontname = "Calibri", part = "all") %>%
    flextable::fontsize(size = 11, part = "all") %>%
    # Column widths
    flextable::width(j = 1, width = 0.8) %>%
    flextable::width(j = 2, width = 0.6) %>%
    flextable::width(j = 3, width = 3.5) %>%
    flextable::width(j = 4, width = 0.6) %>%
    flextable::width(j = 5, width = 0.6) %>%
    # Borders - all width = 1
    flextable::border_remove() %>%
    flextable::hline_top(border = officer::fp_border(width = 1), part = "header") %>%
    flextable::hline_bottom(border = officer::fp_border(width = 1), part = "header") %>%
    flextable::hline_bottom(border = officer::fp_border(width = 1), part = "body")

  # Add separators and zebra striping by quarter/month
  quarter_months <- df %>%
    distinct(quarter, Month) %>%
    mutate(qm = paste(quarter, Month))

  if (nrow(quarter_months) > 1) {
    # Add separators between each quarter/month combination
    for (i in 1:(nrow(quarter_months) - 1)) {
      qm_current <- quarter_months$qm[i]
      last_row <- max(which(paste(df$quarter, df$Month) == qm_current))
      ft <- ft %>%
        flextable::hline(i = last_row, border = officer::fp_border(width = 1))
    }
  }

  # Zebra striping within each quarter/month
  for (i in 1:nrow(quarter_months)) {
    qm <- quarter_months$qm[i]
    qm_rows <- which(paste(df$quarter, df$Month) == qm)
    shade_rows <- qm_rows[seq(2, length(qm_rows), 2)]
    if (length(shade_rows) > 0) {
      ft <- ft %>%
        flextable::bg(i = shade_rows, bg = "#F0F0F0", part = "body")
    }
  }

  # Add table title with bold formatting in Calibri 11
  ft <- ft %>%
    flextable::set_caption(
      caption = as_paragraph(
        as_chunk(table_title,
                 props = officer::fp_text(font.family = "Calibri", font.size = 11, bold = TRUE))
      ),
      style = "Normal",
      align_with_table = FALSE,
      fp_p = officer::fp_par(text.align = "left", padding.bottom = 6)
    )

  # Add footnote if provided (not italicized)
  if (!is.null(footnote)) {
    ft <- ft %>%
      flextable::add_footer_lines(footnote) %>%
      flextable::align(align = "left", part = "footer") %>%
      flextable::font(fontname = "Calibri", part = "footer") %>%
      flextable::fontsize(size = 11, part = "footer")
  }

  ft
}