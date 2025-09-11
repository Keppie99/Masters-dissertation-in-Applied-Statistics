library(readxl)
library(missForest)
library(trend)
library(precintcon)
library(writexl)
library(fmsb)
library(missForest)

#####PRECIPITATION#####
##GariepDam##
GariepDam = read_excel("D:\\Masters\\Precipitation.xlsx",sheet= 23)
colnames(GariepDam) = GariepDam[2,]
GariepDam = GariepDam[-c(1:3),]
GariepDam = data.frame(sapply(GariepDam, as.numeric))
GariepDam$Year = as.factor(GariepDam$Year)
GariepDam = missForest(GariepDam) #MISS_FOREST IMPUTATION
GariepDam = GariepDam$ximp
GariepDam = cbind(GariepDam,"Winter"=rowMeans(GariepDam[c(6:8)]))
GariepDam = cbind(GariepDam,"Spring"=rowMeans(GariepDam[c(9:11)]))
GariepDam = cbind(GariepDam,"Summer"=rowMeans(GariepDam[c(12,13,2)]))
GariepDam = cbind(GariepDam,"Autumn"=rowMeans(GariepDam[c(3:5)]))
GariepDam = cbind(GariepDam,"Annual"=rowMeans(GariepDam[c(2:13)]))
write_xlsx(GariepDam,"D:\\Masters\\Provinces\\FS\\Gariep-Preci.xlsx")
sens.slope(GariepDam$JUN)

Gariep_radar = read_excel("D:\\Masters\\Provinces\\FS\\Gariep-Preci.xlsx",sheet= 2)
radarchart(Gariep_radar[2:18],seg = 3,title = "Gariep Dam")

##BLOEM STAD##
BloemStad = read_excel("D:\\Masters\\Precipitation.xlsx",sheet= 29)
colnames(BloemStad) = BloemStad[2,]
BloemStad = BloemStad[-c(1:3),]
BloemStad = data.frame(sapply(BloemStad, as.numeric))
BloemStad$Year = as.factor(BloemStad$Year)
BloemStad = missForest(BloemStad) #MISS_FOREST IMPUTATION
BloemStad = BloemStad$ximp
BloemStad = cbind(BloemStad,"Winter"=rowMeans(BloemStad[c(6:8)]))
BloemStad = cbind(BloemStad,"Spring"=rowMeans(BloemStad[c(9:11)]))
BloemStad = cbind(BloemStad,"Summer"=rowMeans(BloemStad[c(12,13,2)]))
BloemStad = cbind(BloemStad,"Autumn"=rowMeans(BloemStad[c(3:5)]))
BloemStad = cbind(BloemStad,"Annual"=rowMeans(BloemStad[c(2:13)]))
write_xlsx(BloemStad,"D:\\Masters\\Provinces\\FS\\BloemStad-Preci.xlsx")
sens.slope(BloemStad$Autumn)

Gariep_radar = read_excel("D:\\Masters\\Provinces\\FS\\BloemStad-Preci.xlsx",sheet= 2)
radarchart(Gariep_radar[2:18],seg = 3,title = "Bloem Stad")

##BLOEM WO##
BloemWo = read_excel("D:\\Masters\\Precipitation.xlsx",sheet= 30)
colnames(BloemWo) = BloemWo[2,]
BloemWo = BloemWo[-c(1:3),]
BloemWo = data.frame(sapply(BloemWo, as.numeric))
BloemWo$Year = as.factor(BloemWo$Year)
BloemWo = missForest(BloemWo) #MISS_FOREST IMPUTATION
BloemWo = BloemWo$ximp
BloemWo = cbind(BloemWo,"Winter"=rowMeans(BloemWo[c(6:8)]))
BloemWo = cbind(BloemWo,"Spring"=rowMeans(BloemWo[c(9:11)]))
BloemWo = cbind(BloemWo,"Summer"=rowMeans(BloemWo[c(12,13,2)]))
BloemWo = cbind(BloemWo,"Autumn"=rowMeans(BloemWo[c(3:5)]))
BloemWo = cbind(BloemWo,"Annual"=rowMeans(BloemWo[c(2:13)]))
write_xlsx(BloemWo,"D:\\Masters\\Provinces\\FS\\BloemWo-Preci.xlsx")
sens.slope(BloemWo$Annual)

Gariep_radar = read_excel("D:\\Masters\\Provinces\\FS\\BloemWo-Preci.xlsx",sheet= 2)
radarchart(Gariep_radar[2:18],seg = 3,title = "Bloem Wo")

##FAURESMITH##
Fauresmith = read_excel("D:\\Masters\\Precipitation.xlsx",sheet= 31)
colnames(Fauresmith) = Fauresmith[2,]
Fauresmith = Fauresmith[-c(1:3),]
Fauresmith = data.frame(sapply(Fauresmith, as.numeric))
Fauresmith$Year = as.factor(Fauresmith$Year)
Fauresmith = missForest(Fauresmith) #MISS_FOREST IMPUTATION
Fauresmith = Fauresmith$ximp
Fauresmith = cbind(Fauresmith,"Winter"=rowMeans(Fauresmith[c(6:8)]))
Fauresmith = cbind(Fauresmith,"Spring"=rowMeans(Fauresmith[c(9:11)]))
Fauresmith = cbind(Fauresmith,"Summer"=rowMeans(Fauresmith[c(12,13,2)]))
Fauresmith = cbind(Fauresmith,"Autumn"=rowMeans(Fauresmith[c(3:5)]))
Fauresmith = cbind(Fauresmith,"Annual"=rowMeans(Fauresmith[c(2:13)]))
write_xlsx(Fauresmith,"D:\\Masters\\Provinces\\FS\\Fauresmith-Preci.xlsx")
sens.slope(Fauresmith$Autumn)

Gariep_radar = read_excel("D:\\Masters\\Provinces\\FS\\Fauresmith-Preci.xlsx",sheet= 2)
radarchart(Gariep_radar[2:18],seg = 3,title = "Fauresmith")

##BETHLEHEM##
Bethlehem = read_excel("D:\\Masters\\Precipitation.xlsx",sheet= 32)
colnames(Bethlehem) = Bethlehem[2,]
Bethlehem = Bethlehem[-c(1:3),]
Bethlehem = data.frame(sapply(Bethlehem, as.numeric))
Bethlehem$Year = as.factor(Bethlehem$Year)
Bethlehem = missForest(Bethlehem) #MISS_FOREST IMPUTATION
Bethlehem = Bethlehem$ximp
Bethlehem = cbind(Bethlehem,"Winter"=rowMeans(Bethlehem[c(6:8)]))
Bethlehem = cbind(Bethlehem,"Spring"=rowMeans(Bethlehem[c(9:11)]))
Bethlehem = cbind(Bethlehem,"Summer"=rowMeans(Bethlehem[c(12,13,2)]))
Bethlehem = cbind(Bethlehem,"Autumn"=rowMeans(Bethlehem[c(3:5)]))
Bethlehem = cbind(Bethlehem,"Annual"=rowMeans(Bethlehem[c(2:13)]))
write_xlsx(Bethlehem,"D:\\Masters\\Provinces\\FS\\Bethlehem-Preci.xlsx")
sens.slope(Bethlehem$Spring)

Gariep_radar = read_excel("D:\\Masters\\Provinces\\FS\\Bethlehem-Preci.xlsx",sheet= 2)
radarchart(Gariep_radar[2:18],seg = 3,title = "Bethlehem")

##WELKOM##
Welkom = read_excel("D:\\Masters\\Precipitation.xlsx",sheet= 33)
colnames(Welkom) = Welkom[2,]
Welkom = Welkom[-c(1:3),]
Welkom = data.frame(sapply(Welkom, as.numeric))
Welkom$Year = as.factor(Welkom$Year)
Welkom = missForest(Welkom) #MISS_FOREST IMPUTATION
Welkom = Welkom$ximp
Welkom = cbind(Welkom,"Winter"=rowMeans(Welkom[c(6:8)]))
Welkom = cbind(Welkom,"Spring"=rowMeans(Welkom[c(9:11)]))
Welkom = cbind(Welkom,"Summer"=rowMeans(Welkom[c(12,13,2)]))
Welkom = cbind(Welkom,"Autumn"=rowMeans(Welkom[c(3:5)]))
Welkom = cbind(Welkom,"Annual"=rowMeans(Welkom[c(2:13)]))
write_xlsx(Welkom,"D:\\Masters\\Provinces\\FS\\Welkom-Preci.xlsx")
sens.slope(Welkom$Autumn)

Gariep_radar = read_excel("D:\\Masters\\Provinces\\FS\\Welkom-Preci.xlsx",sheet= 2)
radarchart(Gariep_radar[2:18],seg = 3,title = "Welkom")

##VREDE##
Vrede = read_excel("D:\\Masters\\Precipitation.xlsx",sheet= 34)
colnames(Vrede) = Vrede[2,]
Vrede = Vrede[-c(1:3),]
Vrede = data.frame(sapply(Vrede, as.numeric))
Vrede$Year = as.factor(Vrede$Year)
Vrede = missForest(Vrede) #MISS_FOREST IMPUTATION
Vrede = Vrede$ximp
Vrede = cbind(Vrede,"Winter"=rowMeans(Vrede[c(6:8)]))
Vrede = cbind(Vrede,"Spring"=rowMeans(Vrede[c(9:11)]))
Vrede = cbind(Vrede,"Summer"=rowMeans(Vrede[c(12,13,2)]))
Vrede = cbind(Vrede,"Autumn"=rowMeans(Vrede[c(3:5)]))
Vrede = cbind(Vrede,"Annual"=rowMeans(Vrede[c(2:13)]))
write_xlsx(Vrede,"D:\\Masters\\Provinces\\FS\\Vrede-Preci.xlsx")
pettitt.test(Vrede$Spring)

Gariep_radar = read_excel("D:\\Masters\\Provinces\\FS\\Vrede-Preci.xlsx",sheet= 2)
radarchart(Gariep_radar[2:18],seg = 3,title = "Vrede")

MFI = read_excel("D:\\Masters\\Provinces\\FS\\MFI-FS.xlsx")
MFI = data.frame(t(MFI))
colnames(MFI) = MFI[1,]
MFI = MFI[-1,]
MFI = sapply(MFI,as.numeric)
boxplot(MFI,col="light blue",main="Annual MFI index by gauge station in the Free State province",ylab="MFI index")
abline(h=120, col="#CC3333")

PCI = read_excel("D:\\Masters\\Provinces\\FS\\TCI-FS.xlsx")
boxplot(PCI,col="light blue",main="Annual TCI index by gauge stations in Free State province",ylab="TCI index")
abline(h=15,col="#CC3333")



#####TEMPERATURE#####
##GariepDam##
GariepDam = read_excel("D:\\Masters\\Monthly Temperature Per Station.xlsx",sheet= 1)
colnames(GariepDam) = GariepDam[2,]
GariepDam = GariepDam[-c(1:3),]
GariepDam = data.frame(sapply(GariepDam, as.numeric))
GariepDam$Year = as.factor(GariepDam$Year)
GariepDam = missForest(GariepDam) #MISS_FOREST IMPUTATION
GariepDam = GariepDam$ximp
GariepDam = cbind(GariepDam,"Winter"=rowMeans(GariepDam[c(6:8)]))
GariepDam = cbind(GariepDam,"Spring"=rowMeans(GariepDam[c(9:11)]))
GariepDam = cbind(GariepDam,"Summer"=rowMeans(GariepDam[c(12,13,2)]))
GariepDam = cbind(GariepDam,"Autumn"=rowMeans(GariepDam[c(3:5)]))
GariepDam = cbind(GariepDam,"Annual"=rowMeans(GariepDam[c(2:13)]))
write_xlsx(GariepDam,"D:\\Masters\\Provinces\\FS\\Gariep-Temp.xlsx")
sens.slope(GariepDam$Annual)

Gariep_radar = read_excel("D:\\Masters\\Provinces\\FS\\Gariep-Temp.xlsx",sheet= 2)
radarchart(Gariep_radar[2:18],seg = 3,title = "Gariep Dam")

##BLOEM STAD##
BloemStad = read_excel("D:\\Masters\\Monthly Temperature Per Station.xlsx",sheet= 2)
colnames(BloemStad) = BloemStad[2,]
BloemStad = BloemStad[-c(1:3),]
BloemStad = data.frame(sapply(BloemStad, as.numeric))
BloemStad$Year = as.factor(BloemStad$Year)
BloemStad = missForest(BloemStad) #MISS_FOREST IMPUTATION
BloemStad = BloemStad$ximp
BloemStad = cbind(BloemStad,"Winter"=rowMeans(BloemStad[c(6:8)]))
BloemStad = cbind(BloemStad,"Spring"=rowMeans(BloemStad[c(9:11)]))
BloemStad = cbind(BloemStad,"Summer"=rowMeans(BloemStad[c(12,13,2)]))
BloemStad = cbind(BloemStad,"Autumn"=rowMeans(BloemStad[c(3:5)]))
BloemStad = cbind(BloemStad,"Annual"=rowMeans(BloemStad[c(2:13)]))
write_xlsx(BloemStad,"D:\\Masters\\Provinces\\FS\\BloemStad-Temp.xlsx")
sens.slope(BloemStad$Autumn)

Gariep_radar = read_excel("D:\\Masters\\Provinces\\FS\\BloemStad-Temp.xlsx",sheet= 2)
radarchart(Gariep_radar[2:18],seg = 3,title = "Bloem Stad")

##BLOEM WO##
BloemWo = read_excel("D:\\Masters\\Monthly Temperature Per Station.xlsx",sheet= 3)
colnames(BloemWo) = BloemWo[2,]
BloemWo = BloemWo[-c(1:3),]
BloemWo = data.frame(sapply(BloemWo, as.numeric))
BloemWo$Year = as.factor(BloemWo$Year)
BloemWo = missForest(BloemWo) #MISS_FOREST IMPUTATION
BloemWo = BloemWo$ximp
BloemWo = cbind(BloemWo,"Winter"=rowMeans(BloemWo[c(6:8)]))
BloemWo = cbind(BloemWo,"Spring"=rowMeans(BloemWo[c(9:11)]))
BloemWo = cbind(BloemWo,"Summer"=rowMeans(BloemWo[c(12,13,2)]))
BloemWo = cbind(BloemWo,"Autumn"=rowMeans(BloemWo[c(3:5)]))
BloemWo = cbind(BloemWo,"Annual"=rowMeans(BloemWo[c(2:13)]))
write_xlsx(BloemWo,"D:\\Masters\\Provinces\\FS\\BloemWo-Temp.xlsx")
sens.slope(BloemWo$Annual)

Gariep_radar = read_excel("D:\\Masters\\Provinces\\FS\\BloemWo-Temp.xlsx",sheet= 2)
radarchart(Gariep_radar[2:18],seg = 3,title = "Bloem Wo")

##FAURESMITH##
Fauresmith = read_excel("D:\\Masters\\Monthly Temperature Per Station.xlsx",sheet= 4)
colnames(Fauresmith) = Fauresmith[2,]
Fauresmith = Fauresmith[-c(1:3),]
Fauresmith = data.frame(sapply(Fauresmith, as.numeric))
Fauresmith$Year = as.factor(Fauresmith$Year)
Fauresmith = missForest(Fauresmith) #MISS_FOREST IMPUTATION
Fauresmith = Fauresmith$ximp
Fauresmith = cbind(Fauresmith,"Winter"=rowMeans(Fauresmith[c(6:8)]))
Fauresmith = cbind(Fauresmith,"Spring"=rowMeans(Fauresmith[c(9:11)]))
Fauresmith = cbind(Fauresmith,"Summer"=rowMeans(Fauresmith[c(12,13,2)]))
Fauresmith = cbind(Fauresmith,"Autumn"=rowMeans(Fauresmith[c(3:5)]))
Fauresmith = cbind(Fauresmith,"Annual"=rowMeans(Fauresmith[c(2:13)]))
write_xlsx(Fauresmith,"D:\\Masters\\Provinces\\FS\\Fauresmith-Temp.xlsx")
sens.slope(Fauresmith$Annual)

Gariep_radar = read_excel("D:\\Masters\\Provinces\\FS\\Fauresmith-Temp.xlsx",sheet= 2)
radarchart(Gariep_radar[2:18],seg = 3,title = "Fauresmith")

##BETHLEHEM##
Bethlehem = read_excel("D:\\Masters\\Monthly Temperature Per Station.xlsx",sheet= 5)
colnames(Bethlehem) = Bethlehem[2,]
Bethlehem = Bethlehem[-c(1:3),]
Bethlehem = data.frame(sapply(Bethlehem, as.numeric))
Bethlehem$Year = as.factor(Bethlehem$Year)
Bethlehem = missForest(Bethlehem) #MISS_FOREST IMPUTATION
Bethlehem = Bethlehem$ximp
Bethlehem = cbind(Bethlehem,"Winter"=rowMeans(Bethlehem[c(6:8)]))
Bethlehem = cbind(Bethlehem,"Spring"=rowMeans(Bethlehem[c(9:11)]))
Bethlehem = cbind(Bethlehem,"Summer"=rowMeans(Bethlehem[c(12,13,2)]))
Bethlehem = cbind(Bethlehem,"Autumn"=rowMeans(Bethlehem[c(3:5)]))
Bethlehem = cbind(Bethlehem,"Annual"=rowMeans(Bethlehem[c(2:13)]))
write_xlsx(Bethlehem,"D:\\Masters\\Provinces\\FS\\Bethlehem-Temp.xlsx")
sens.slope(Bethlehem$Autumn)

Gariep_radar = read_excel("D:\\Masters\\Provinces\\FS\\Bethlehem-Temp.xlsx",sheet= 2)
radarchart(Gariep_radar[2:18],seg = 3,title = "Bethlehem")

##WELKOM##
Welkom = read_excel("D:\\Masters\\Monthly Temperature Per Station.xlsx",sheet= 6)
colnames(Welkom) = Welkom[2,]
Welkom = Welkom[-c(1:3),]
Welkom = data.frame(sapply(Welkom, as.numeric))
Welkom$Year = as.factor(Welkom$Year)
Welkom = missForest(Welkom) #MISS_FOREST IMPUTATION
Welkom = Welkom$ximp
Welkom = cbind(Welkom,"Winter"=rowMeans(Welkom[c(6:8)]))
Welkom = cbind(Welkom,"Spring"=rowMeans(Welkom[c(9:11)]))
Welkom = cbind(Welkom,"Summer"=rowMeans(Welkom[c(12,13,2)]))
Welkom = cbind(Welkom,"Autumn"=rowMeans(Welkom[c(3:5)]))
Welkom = cbind(Welkom,"Annual"=rowMeans(Welkom[c(2:13)]))
write_xlsx(Welkom,"D:\\Masters\\Provinces\\FS\\Welkom-Temp.xlsx")
sens.slope(Welkom$Autumn)

Gariep_radar = read_excel("D:\\Masters\\Provinces\\FS\\Vrede-Temp.xlsx",sheet= 2)
radarchart(Gariep_radar[2:18],seg = 3,title = "Vrede")

##VREDE##
Vrede = read_excel("D:\\Masters\\Monthly Temperature Per Station.xlsx",sheet= 7)
colnames(Vrede) = Vrede[2,]
Vrede = Vrede[-c(1:3),]
Vrede = data.frame(sapply(Vrede, as.numeric))
Vrede$Year = as.factor(Vrede$Year)
Vrede = missForest(Vrede) #MISS_FOREST IMPUTATION
Vrede = Vrede$ximp
Vrede = cbind(Vrede,"Winter"=rowMeans(Vrede[c(6:8)]))
Vrede = cbind(Vrede,"Spring"=rowMeans(Vrede[c(9:11)]))
Vrede = cbind(Vrede,"Summer"=rowMeans(Vrede[c(12,13,2)]))
Vrede = cbind(Vrede,"Autumn"=rowMeans(Vrede[c(3:5)]))
Vrede = cbind(Vrede,"Annual"=rowMeans(Vrede[c(2:13)]))
write_xlsx(Vrede,"D:\\Masters\\Provinces\\FS\\Vrede-Temp.xlsx")
sens.slope(Vrede$Autumn)
