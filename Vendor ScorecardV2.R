
#Loading required packages
library(readxl)
library(sqldf)
options(scipen = 999)

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
                      t3.mnt_no as Month,
                      count(*) as SKUCount,
                      1-avg(t2.`Cost Price`)/avg(t2.MRP) as TR
               from UCM1 as t1
               left join UCM2 as t2
                      on t1.`Seller SKU on Channel` = t2.`Product Code`
               left join Calendar_Mod as t3
                      on t1. Createdmnt <= t3.mnt_no
                     and t3.mnt_no <= '2023-03'
               left join FacilityMap as t4
                      on t2.Facility_code = t4.`Facility Code`
               group by 1,2")

UCM1 <- NULL
UCM2 <- NULL


#Sales

uc_raw <- read_excel("C:/Users/harsh/Desktop/Vaaree/Sales/UC_Sales.xlsx")

#Retaining only the required columns & renaming 

old_column_names <- c("Sale Order Item Code","Display Order Code","Reverse Pickup Code","COD","Category","Shipping Address Pincode","Item SKU Code","Channel Product Id","MRP","Selling Price","Cost Price","Discount","Voucher Code","Packet Number","Order Date as dd/mm/yyyy hh:MM:ss","Sale Order Status","Sale Order Item Status","Cancellation Reason","Shipping provider","Shipping Courier","Shipping Package Creation Date","Shipping Package Status Code","Delivery Time","Tracking Number","Dispatch Date","Facility","Return Reason")
new_column_names <- c("SaleOrderItemCode","DisplayOrderCode","ReversePickupCode","COD","Category","ShippingAddressPincode","ItemSKUCode","ChannelProductId","MRP","SellingPrice","CostPrice","Discount","VoucherCode","UnitsSold","OrderDate","SaleOrderStatus","SaleOrderItemStatus","CancellationReason","ShippingProvider","ShippingCourier","ShippingPackageCreationDate","ShippingPackageStatusCode","DeliveryTime","AWB","DispatchDate","Facility","ReturnReason")

uc_raw <- uc_raw[, old_column_names]
colnames(uc_raw) <- new_column_names


uc_raw$OrderMonth <- format(as.Date(uc_raw$OrderDate),"%Y-%m")
uc_raw$OrderDate <- as.character(as.Date(uc_raw$OrderDate))
uc_raw$ShippingPackageCreationDate <- as.character(as.Date(uc_raw$ShippingPackageCreationDate))
uc_raw$DeliveryTime <- as.character(as.Date(uc_raw$DeliveryTime))
uc_raw$DispatchDate <- as.character(as.Date(uc_raw$DispatchDate))

temp <- sqldf("select DisplayOrderCode,
                      Facility,
                      OrderDate,
                      DispatchDate,
                      count(*)-1 as DaysToShip
               from
               (select DisplayOrderCode,
                       Facility,
                       OrderDate,
                       DispatchDate
               from uc_raw
               where DispatchDate is not null
               group by 1,2,3,4) as t1
               left join Calendar as t2
                      on t2.dt between t1.OrderDate and t1.DispatchDate
                     and t2.day != 7 and t2.holiday_flag = 0
               group by 1,2,3,4")

ret_temp <- sqldf("select `Order Number` as OrderNo,
                           max(case when `VV Reason` = 'Defective/Damaged' then 1 else 0 end) Defective_Flag,
                           max(case when `VV Reason` = 'Customer Reasons' then 1 else 0 end) Customer_Flag,
                           max(case when `VV Reason` = 'Wrong Dispatch' then 1 else 0 end) WrongDispatch_Flag
                   from Returns
                   group by 1")

Fin2 <- sqldf("select t1.Facility,
                      t1.OrderMonth,
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
                      (sum(t1.SellingPrice) + sum(t1.Discount)) as LP,
                      sum(t1.CostPrice) as Cost
               from uc_raw as t1
               left join temp as t2
                      on t1.DisplayOrderCode = t2.DisplayOrderCode
                     and t1.Facility = t2.Facility
               left join ret_temp as t3
                      on t1.DisplayOrderCode = t3.OrderNo
               group by 1,2")

JudgeMe$ReviewMonth <- substring(JudgeMe$review_date, 1,7)
FacilityMap$`Facility Code`<- as.character(FacilityMap$`Facility Code`)

Fin3 <- sqldf("select t5.`Vendor Name`,
                      t4.Facility_code,
                      t3.mnt_no,
                      avg(t1.rating) as Rating,
                      count(*) as NoRating
               from JudgeMe as t1
               left join Matrixify as t2
                      on t1.product_id = t2.ID
               left join Calendar_Mod as t3
                      on t1.ReviewMonth <= t3.mnt_no
                     and t3.mnt_no <= '2023-03'
               left join uc_item_master as t4
                      on t2.`Variant SKU`= t4.`Scan Identifier`
               left join FacilityMap as t5
                      on t5.`Facility Code` = t4.Facility_code
               where t5.`Vendor Name` is not null
               group by 1,2,3")


Fin4 <- sqldf("select  Facility,
                       avg(Inventory) as AVG_Stock,
                       count(case when Inventory = 0 then 1 end) as OOS_Items,
                       count(*) as Total_Items
               from Inventory
               group by 1")


FinFin <- sqldf("select t1.*,
                        coalesce(t2.OrderCount,0) as OrderCount,
                        coalesce(t2.AvgDaysToShip,0) as AvgDaystoShip,
                        coalesce(t2.DefectiveOrders,0) as DefectiveOrders,
                        coalesce(t2.WrongDispatchOrders,0) as WrongDispatchOrders,
                        coalesce(t2.RTO_Orders,0) as RTO_Orders
                 from Fin1 as t1
                 left join Fin2 as t2
                        on t1.Brand = t2.Facility
                       and t2.OrderMonth = t1.Month")

#Making the final output

base <- sqldf("select Brand from Fin1 where Brand is not null group by 1
               union
               select Facility from Fin2 where Facility is not null group by 1
               union 
               select `Vendor Name` from Fin3 where `Vendor Name` is not null group by 1
               union 
               select Facility from Fin4 where Facility is not null group by 1")


finale <- sqldf("select t1.Brand,
                        t5.AVG_Stock,
                        t3.OrderCount as `Orders/Month`,
                        t3.AvgDaysToShip,
                        t5.OOS_Items*1.0/t5.Total_Items*1.0 as OOS_Perc,
                        (t3.DefectiveOrders + t3.CustomersOrders + t3.WrongDispatchOrders)*1.0/t3.OrderCount*1.0 as `Return%`,
                        t4.Rating,
                        t3.LP,
                        t2.SKUCount,
                        t2.TR as TakeRate,
                        case when t2.SKUCount <50 then 0
                             when t2.SKUCount <100 then 1
                             when t2.SKUCount >=100 then 2
                             else 0
                        end as SKUCountScore,
                        case when t5.AVG_Stock <10 then 0
                             when t5.AVG_Stock <50 then 1
                             when t5.AVG_Stock >=50 then 2
                             else 0
                        end as InvtScore,
                        case when t3.OrderCount <20 then 0
                             when t3.OrderCount <50 then 1
                             when t3.OrderCount <200 then 2
                             when t3.OrderCount >=200 then 3
                             else 0
                        end as OrderScore,
                        case when t3.AvgDaysToShip <=1 then 2
                             when t3.AvgDaysToShip <=2 then 1
                             when t3.AvgDaysToShip >2 then 0
                             else 0
                        end as DispScore,
                        case when t5.OOS_Items/t5.Total_Items <0.05 then 2
                             when t5.OOS_Items/t5.Total_Items <0.25 then 1
                             when t5.OOS_Items/t5.Total_Items >=0.25 then 0
                             else 0
                        end as OOSScore,
                        case when (t3.DefectiveOrders + t3.WrongDispatchOrders)/t3.OrderCount <0.05 then 2
                             when (t3.DefectiveOrders + t3.WrongDispatchOrders)/t3.OrderCount <0.1 then 1
                             when (t3.DefectiveOrders + t3.WrongDispatchOrders)/t3.OrderCount >=0.1 then 0
                             else 0
                        end as RetScore,
                        case when t4.Rating <3 then 0
                             when t4.Rating <=4 then 1
                             when t4.Rating >4 then 2
                             else 0
                        end as RatScore
                 from base as t1
                 left join Fin1 as t2
                        on t1.Brand = t2.Brand
                       and t2.Month = '2023-03'
                 left join Fin2 as t3
                        on t1.Brand = t3.Facility
                       and t3.OrderMonth = '2023-03'
                 left join Fin3 as t4
                        on t1.Brand = t4.`Vendor Name`
                       and t4.mnt_no = '2023-03'
                 left join Fin4 as t5
                        on t1.Brand = t5.Facility")


write.table(finale,"clipboard-32768", sep="\t", col.names=TRUE,row.names = FALSE)



