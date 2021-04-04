
context("Testing 'topN' and 'bottomN' conditions in conditionalFormatting")
TNBN_test_data <- data.frame(col1 = 1:10,
                             col2 = 1:10,
                             col3 = seq(10, 100, 10),
                             col4 = seq(10, 100, 10),
                             col5 = 1:10,
                             col6 = 1:10)

bg_blue <- createStyle(bgFill = "skyblue")

wb <- createWorkbook()
sht <- "TopN_BottomN_TEST"
addWorksheet(wb, sht)
writeData(wb, sht, TNBN_test_data)
conditionalFormatting(wb, sht, cols = 1, rows = 2:11, type = "topN",    rank = 3,  style = bg_blue, percent = FALSE)
conditionalFormatting(wb, sht, cols = 2, rows = 2:11, type = "bottomN", rank = 3,  style = bg_blue, percent = FALSE)
conditionalFormatting(wb, sht, cols = 3, rows = 2:11, type = "topN",    rank = 50, style = bg_blue, percent = TRUE)
conditionalFormatting(wb, sht, cols = 4, rows = 2:11, type = "bottomN", rank = 50, style = bg_blue, percent = TRUE)
conditionalFormatting(wb, sht, cols = 5, rows = 2:11, type = "topN",    rank = 3,  style = bg_blue)
conditionalFormatting(wb, sht, cols = 6, rows = 2:11, type = "bottomN", rank = 3,  style = bg_blue)

test_that("Number of conditionalFormatting rules added equal to 6", {
  expect_equal(object = length(wb$worksheets[[1]]$conditionalFormatting), expected = 6)
})

test_that("topN conditions do not have the 'bottom' argument", {
  expect_false(object = grepl(paste('bottom'), wb$worksheets[[1]]$conditionalFormatting[1]))
  expect_false(object = grepl(paste('bottom'), wb$worksheets[[1]]$conditionalFormatting[3]))
})

test_that("bottomN conditions have the 'bottom' argument set to '1'", {
  expect_true(object = grepl(paste('bottom="1"'), wb$worksheets[[1]]$conditionalFormatting[2]))
  expect_true(object = grepl(paste('bottom="1"'), wb$worksheets[[1]]$conditionalFormatting[4]))
})

test_that("topN/bottomN rank conditions have the 'percent=FALSE' argument set to '0'", {
  expect_true(object = grepl(paste('percent="0"'), wb$worksheets[[1]]$conditionalFormatting[1]))
  expect_true(object = grepl(paste('percent="0"'), wb$worksheets[[1]]$conditionalFormatting[2]))
})

test_that("topN/bottomN rank conditions do not have the 'percent' argument is set to 'NULL'", {
  expect_true(object = grepl(paste('percent="NULL"'), wb$worksheets[[1]]$conditionalFormatting[5]))
  expect_true(object = grepl(paste('percent="NULL"'), wb$worksheets[[1]]$conditionalFormatting[6]))
})

test_that("topN/bottomN percent conditions have the 'percent' argument set to '1'", {
  expect_true(object = grepl(paste('percent="1"'), wb$worksheets[[1]]$conditionalFormatting[3]))
  expect_true(object = grepl(paste('percent="1"'), wb$worksheets[[1]]$conditionalFormatting[4]))
})

test_that("topN/bottomN conditions correspond to 'top10' type", {
  expect_true(object = grepl(paste('type="top10"'), wb$worksheets[[1]]$conditionalFormatting[1]))
  expect_true(object = grepl(paste('type="top10"'), wb$worksheets[[1]]$conditionalFormatting[2]))
  expect_true(object = grepl(paste('type="top10"'), wb$worksheets[[1]]$conditionalFormatting[3]))
  expect_true(object = grepl(paste('type="top10"'), wb$worksheets[[1]]$conditionalFormatting[4]))
  expect_true(object = grepl(paste('type="top10"'), wb$worksheets[[1]]$conditionalFormatting[5]))
  expect_true(object = grepl(paste('type="top10"'), wb$worksheets[[1]]$conditionalFormatting[6]))
})


context("Testing iconSet condition in conditionalFormatting")
set.seed(42)
## iconSet displays icons inside cells based on cell value
wb <- createWorkbook()
addWorksheet(wb, "iconSet")
df <- data.frame(col1 = rep(c(-1,0,1), 3), 
                 col2 = rep(c(-1,0,1), 3), 
                 col3 = rep(c(-1,0,1), 3),
                 col4 = sample(1:150, 9),
                 col5 = sample(100:200, 9)
                 )
writeData(wb, "iconSet", df, colNames = FALSE) ## write data.frame

conditionalFormatting(wb, "iconSet",
 cols = 1, rows = 1:nrow(df),
 type = "iconSet",
 rule = "iconSet",
 limits = list(type = c("percent","percent","percent"), val = c("0","33","67"))
)
conditionalFormatting(wb, "iconSet",
 cols = 2, rows = 1:nrow(df),
 type = "iconSet",
 rule = "3Symbols", # Red-Amber-Green Circles with cross, warning and tick inside
 limits = list(type = c("percent","percent","percent"), val = c("0","33","67")),
 showValue = FALSE # show Icons only
)
conditionalFormatting(wb, "iconSet",
 cols = 3, rows = 1:nrow(df),
 type = "iconSet",
 rule = "3Symbols2", # Red-Amber-Green cross, warning and tick symbols
 limits = list(type = c("percent","percent","percent"), val = c("0","33","67")),
 reverse = TRUE
)
conditionalFormatting(wb, "iconSet",
 cols = 4, rows = 1:nrow(df),
 type = "iconSet",
 rule = "3Flags", 
 limits = list(type = c("num","num","num"), val = c("50","99","150"))
)
conditionalFormatting(wb, "iconSet",
 cols = 5, rows = 1:nrow(df),
 type = "iconSet",
 rule = "3Arrows", # Red-Amber-Green cross, warning and tick symbols
 limits = list(type = c("num","num","num"), val = c("150","175","200"))
)

saveWorkbook(wb, "C:/Users/Vishal/Desktop/testing.xlsx", overwrite = TRUE)
openXL("C:/Users/Vishal/Desktop/testing.xlsx")

test_that("rule='iconSet' does not add a iconSet attribute to the XML", {
  expect_false(grepl('iconSet=', wb$worksheets[[1]]$conditionalFormatting[1]))
})

test_that("rule !='iconSet' adds a iconSet attribute to the XML", {
  expect_true(grepl('iconSet=', wb$worksheets[[1]]$conditionalFormatting[2]))
  expect_true(grepl('iconSet=', wb$worksheets[[1]]$conditionalFormatting[3]))
  expect_true(grepl('iconSet=', wb$worksheets[[1]]$conditionalFormatting[4]))
  expect_true(grepl('iconSet=', wb$worksheets[[1]]$conditionalFormatting[5]))
})





## rule is one of following names of the icons sets: 
## "iconSet","3Arrows","3ArrowsGray","3Signs","3TrafficLights2",
## "3Symbols","3Symbols2","3Flags".
## if rule is "iconSet", default icon set of Red-Amber-Green Circles is used.
## limits must be a list with 2 named character vectors 'type' and 'val'
## 'type' must have length 3 with elements being either 'num' or 'percent'
## 'val' must have same length as 'type' with values in ascending order.







