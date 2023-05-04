
#Loading required packages
library(readxl)
library(sqldf)


#Importing required excel files
UCM1 <- read_excel("C:/Users/harsh/Desktop/Vaaree/Sales/UC_Item_Master.xlsx")
UCM2 <- read_excel("C:/Users/harsh/Desktop/Vaaree/Sales/UC_Item_Master_New.xlsx")
Calendar <- read_excel("C:/Users/harsh/Desktop/Vaaree/Sales/Calendar.xlsx")
Returns <- read_excel("C:/Users/harsh/Desktop/Vaaree/Sales/Returns.xlsx")
Matrixify <- read_excel("C:/Users/harsh/Desktop/Vaaree/Sales/Matrixify_B.xlsx")
JudgeMe <- read.csv("C:/Users/harsh/Desktop/Vaaree/Sales/JudgeMe_Reviews.csv")
Inventory <- read_excel("C:/Users/harsh/Desktop/Vaaree/Sales/UC_Invt_2022_11_15.xlsx")
uc_item_master <- read_excel("C:/Users/harsh/Desktop/Vaaree/Sales/UC_Item_Master_New.xlsx")
FacilityMap <- read_excel("C:/Users/harsh/Desktop/Vaaree/Sales/Facility Map.xlsx")


#Processing the shopify file
UCM1$CreatedDT <- as.character(as.Date(UCM1$Created))
UCM1$CreatedMnt <- format(as.Date(UCM1$Created),"%Y-%m")

Calendar_Mod <- sqldf("select mnt_no from Calendar group by 1")

Fin1 <- sqldf("select t4.`Vendor Name` as Brand,
                      t2.`Category Code` as Category,
                      t3.mnt_no as Month,
                      count(*) as SKUCount
               from UCM1 as t1
               left join UCM2 as t2
                      on t1.`Seller SKU on Channel` = t2.`Product Code`
               left join Calendar_Mod as t3
                      on t1. Createdmnt <= t3.mnt_no
                     and t3.mnt_no <= '2023-03'
               left join FacilityMap as t4
                      on t2.Facility_code = t4.`Facility Code`
               group by 1,2,3")

UCM1 <- NULL
UCM2 <- NULL


#Sales

uc_raw <- read_excel("C:/Users/harsh/Desktop/Vaaree/Sales/UC_Sales.xlsx")

uc_raw <- uc_raw[,c(1:3,9,11,23,35,36,46,48,49,67,72,73,76,80:83,86,87,92:95,97)]
colnames(uc_raw) <- c("SaleOrderItemCode","DisplayOrderCode","ReversePickupCode","COD","Category","ShippingAddressPincode","ItemSKUCode","ChannelProductId","MRP","SellingPrice","CostPrice","VoucherCode","UnitsSold","OrderDate","SaleOrderStatus","SaleOrderItemStatus","CancellationReason","ShippingProvider","ShippingCourier","ShippingPackageCreationDate","ShippingPackageStatusCode","DeliveryTime","AWB","DispatchDate","Facility","ReturnReason")

uc_raw$OrderMonth <- format(as.Date(uc_raw$OrderDate),"%Y-%m")
uc_raw$OrderDate <- as.character(as.Date(uc_raw$OrderDate))
uc_raw$ShippingPackageCreationDate <- as.character(as.Date(uc_raw$ShippingPackageCreationDate))
uc_raw$DeliveryTime <- as.character(as.Date(uc_raw$DeliveryTime))
uc_raw$DispatchDate <- as.character(as.Date(uc_raw$DispatchDate))


temp <- sqldf("select DisplayOrderCode,
                      OrderDate,
                      DispatchDate,
                      count(*)-1 as DaysToShip
               from
               (select DisplayOrderCode,
                       OrderDate,
                       DispatchDate
               from uc_raw
               where DispatchDate is not null
               group by 1,2,3) as t1
               left join Calendar as t2
                      on t2.dt between t1.OrderDate and t1.DispatchDate
                     and t2.day != 7 and t2.holiday_flag = 0
               group by 1,2,3")

ret_temp <- sqldf("select `Order Number` as OrderNo,
                           max(case when `VV Reason` = 'Defective/Damaged' then 1 else 0 end) Defective_Flag,
                           max(case when `VV Reason` = 'Customer Reasons' then 1 else 0 end) Customer_Flag,
                           max(case when `VV Reason` = 'Wrong Dispatch' then 1 else 0 end) WrongDispatch_Flag
                   from Returns
                   group by 1")

cat_temp <- sqldf("select Category,
                          OrderMonth,
                          count(distinct `DisplayOrderCode`) as CatOrderCount,
                          sum(MRP) as CatLP
                   from uc_raw
                   where DispatchDate is not null
                   group by 1,2")

Fin2 <- sqldf("select t1.Facility,
                      t1.OrderMonth,
                      t1.Category,
                      t4.CatOrderCount,
                      t4.CatLP,
                      count(distinct t1.`DisplayOrderCode`) as OrderCount,
                      avg(DaysToShip) as AvgDaysToShip,
                      count(distinct case when t3.Defective_Flag = 1 then t3.OrderNo end) as DefectiveOrders,
                      count(distinct case when t3.Customer_Flag = 1 then t3.OrderNo end) as CustomersOrders,
                      count(distinct case when t3.WrongDispatch_Flag = 1 then t3.OrderNo end) as WrongDispatchOrders,
                      count(distinct case when t1.ReversePickupCode is null 
                                           and t1.SaleOrderStatus = 'COMPLETE'
                                           and t1.SaleOrderItemStatus in ('DISPATCHED','DELIVERED') 
                                           and t1.ShippingPackageStatusCode in ('RETURNED','RETURN_EXPECTED')
                                      then t1.`DisplayOrderCode` end) as RTO_Orders,
                      sum(t1.MRP) as LP
               from uc_raw as t1
               left join temp as t2
                      on t1.DisplayOrderCode = t2.DisplayOrderCode
               left join ret_temp as t3
                      on t1.DisplayOrderCode = t3.OrderNo
               left join cat_temp as t4
                      on t4.OrderMonth = t1.OrderMonth
                     and t4.Category = t1.Category
               group by 1,2,3,4,5")

JudgeMe$ReviewMonth <- substring(JudgeMe$review_date, 1,7)

Fin3 <- sqldf("select t4.Brand,
                      t3.mnt_no,
                      t4.`Category Code`,
                      avg(t1.rating) as Rating,
                      count(*) as NoRating
               from JudgeMe as t1
               left join Matrixify as t2
                      on t1.product_id = t2.ID
               left join Calendar_Mod as t3
                      on t1.ReviewMonth <= t3.mnt_no
                     and t3.mnt_no <= '2023-03'
               left join uc_item_master as t4
                      on t2.`Variant SKU`= t4.`Product Code`
               where t4.Brand is not null
               group by 1,2,3")


Fin4 <- sqldf("select  Facility,
                       avg(Inventory) as AVG_Stock,
                       count(case when Inventory = 0 then 1 end) as OOS_Items,
                       count(*) as Total_Items
               from Inventory
               group by 1")


FinFin <- sqldf("select t1.*,
                        coalesce(t2.CatOrderCount,0) as CatOrderCount,
                        coalesce(t2.OrderCount,0) as OrderCount,
                        coalesce(t2.AvgDaysToShip,0) as AvgDaystoShip,
                        coalesce(t2.DefectiveOrders,0) as DefectiveOrders,
                        coalesce(t2.WrongDispatchOrders,0) as WrongDispatchOrders,
                        coalesce(t2.RTO_Orders,0) as RTO_Orders
                 from Fin1 as t1
                 left join Fin2 as t2
                        on t1.Brand = t2.Facility
                       and t2.OrderMonth = t1.Month
                       and t1.Category = t2.Category")

write.table(Fin2,"clipboard-32768", sep="\t", col.names=TRUE,row.names = FALSE)

