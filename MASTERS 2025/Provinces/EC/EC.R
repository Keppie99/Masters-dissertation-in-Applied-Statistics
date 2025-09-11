library(readxl)
library(missForest)
library(trend)
library(precintcon)
library(writexl)
library(fmsb)

##TSITSIKAMMA##
tsitsikamma = read_excel("D:\\Masters\\Temperature.xlsx",sheet= 1)
colnames(tsitsikamma) = tsitsikamma[2,]
tsitsikamma = tsitsikamma[-c(1:3),]
tsitsikamma = data.frame(sapply(tsitsikamma, as.numeric))
tsitsikamma$Year = as.factor(tsitsikamma$Year)
tsitsikamma = missForest(tsitsikamma) #MISS_FOREST IMPUTATION
tsitsikamma = tsitsikamma$ximp
tsitsikamma = cbind(tsitsikamma,"Winter"=rowMeans(tsitsikamma[c(6:8)]))
tsitsikamma = cbind(tsitsikamma,"Spring"=rowMeans(tsitsikamma[c(9:11)]))
tsitsikamma = cbind(tsitsikamma,"Summer"=rowMeans(tsitsikamma[c(12,13,2)]))
tsitsikamma = cbind(tsitsikamma,"Autumn"=rowMeans(tsitsikamma[c(3:5)]))
tsitsikamma = cbind(tsitsikamma,"Annual"=rowMeans(tsitsikamma[c(2:13)]))
write_xlsx(tsitsikamma,"D:\\Masters\\Provinces\\EC\\Tsitsi-Temp.xlsx")
mk.test(tsitsikamma$JUL)
tsitsikamma_radar = read_excel("D:\\Masters\\Provinces\\EC\\Tsitsi-Temp.xlsx",sheet= 2)
tsitsikamma = sapply(tsitsikamma_radar, as.numeric)
radarchart(tsitsikamma_radar[2:18],seg = 3,title = "Tsitsikamma")

##CAPE FRANCIS##
CFrancis = read_excel("D:\\Masters\\Temperature.xlsx",sheet= 2)
colnames(CFrancis) = CFrancis[2,]
CFrancis = CFrancis[-c(1:3),]
CFrancis = data.frame(sapply(CFrancis, as.numeric))
CFrancis$Year = as.factor(CFrancis$Year)
CFrancis = missForest(CFrancis) #MISS_FOREST IMPUTATION
CFrancis <- CFrancis$ximp
CFrancis = cbind(CFrancis,"Winter"=rowMeans(CFrancis[c(6:8)]))
CFrancis = cbind(CFrancis,"Spring"=rowMeans(CFrancis[c(9:11)]))
CFrancis = cbind(CFrancis,"Summer"=rowMeans(CFrancis[c(12,13,2)]))
CFrancis = cbind(CFrancis,"Autumn"=rowMeans(CFrancis[c(3:5)]))
CFrancis = cbind(CFrancis,"Annual"=rowMeans(CFrancis[c(2:13)]))
write_xlsx(CFrancis,"D:\\Masters\\Provinces\\EC\\CFrancis-Temp.xlsx")
mk.test(CFrancis$Annual)
CFrancis_radar = read_excel("D:\\Masters\\Provinces\\EC\\CFrancis-Temp.xlsx",sheet= 2)
radarchart(CFrancis_radar[2:18],seg = 3,title = "Cape Francis")
##EAST LONDON##
ELondon = read_excel("D:\\Masters\\Temperature.xlsx",sheet= 3)
colnames(ELondon) = ELondon[2,]
ELondon = ELondon[-c(1:3),]
ELondon = data.frame(sapply(ELondon, as.numeric))
ELondon$Year = as.factor(ELondon$Year)
ELondon = missForest(ELondon) #MISS_FOREST IMPUTATION
ELondon <- ELondon$ximp
ELondon = cbind(ELondon,"Winter"=rowMeans(ELondon[c(6:8)]))
ELondon = cbind(ELondon,"Spring"=rowMeans(ELondon[c(9:11)]))
ELondon = cbind(ELondon,"Summer"=rowMeans(ELondon[c(12,13,2)]))
ELondon = cbind(ELondon,"Autumn"=rowMeans(ELondon[c(3:5)]))
ELondon = cbind(ELondon,"Annual"=rowMeans(ELondon[c(2:13)]))
write_xlsx(ELondon,"D:\\Masters\\Provinces\\EC\\ELondo-Temp.xlsx")

ELondon_radar = read_excel("D:\\Masters\\Provinces\\EC\\ELondo-Temp.xlsx",sheet= 2)
radarchart(CFrancis_radar[2:18],seg = 3,title = "East London")

mk.test(ELondon$Annual)


TCI = read_excel("D:\\Masters\\Provinces\\EC\\TCI.xlsx")
boxplot(TCI,col="light blue",main="Annual TCI index by gauge stations in Eastern Cape province",ylab="TCI index")




###PRECIPITATION###
##TSITSIKAMMA##
tsitsikamma = read_excel("D:\\Masters\\Precipitation.xlsx",sheet= 10)
colnames(tsitsikamma) = tsitsikamma[2,]
tsitsikamma = tsitsikamma[-c(1:3),]
tsitsikamma = data.frame(sapply(tsitsikamma, as.numeric))
tsitsikamma$Year = as.factor(tsitsikamma$Year)
tsitsikamma = missForest(tsitsikamma) #MISS_FOREST IMPUTATION
tsitsikamma = tsitsikamma$ximp
tsitsikamma = cbind(tsitsikamma,"Winter"=rowMeans(tsitsikamma[c(6:8)]))
tsitsikamma = cbind(tsitsikamma,"Spring"=rowMeans(tsitsikamma[c(9:11)]))
tsitsikamma = cbind(tsitsikamma,"Summer"=rowMeans(tsitsikamma[c(12,13,2)]))
tsitsikamma = cbind(tsitsikamma,"Autumn"=rowMeans(tsitsikamma[c(3:5)]))
tsitsikamma = cbind(tsitsikamma,"Annual"=rowMeans(tsitsikamma[c(2:13)]))
write_xlsx(tsitsikamma,"D:\\Masters\\Provinces\\EC\\Tsitsi-Prec.xlsx")

tsitsikamma_radar = read_excel("D:\\Masters\\Provinces\\EC\\Tsitsi-Prec.xlsx",sheet= 2)
radarchart(tsitsikamma_radar[2:18],seg=3,title = "Tsitsikamma")

mk.test(tsitsikamma$Annual)

sens.slope(tsitsikamma$Annual)


##CAPE FRANCIS##
CFrancis = read_excel("D:\\Masters\\Precipitation.xlsx",sheet= 11)
colnames(CFrancis) = CFrancis[2,]
CFrancis = CFrancis[-c(1:3),]
CFrancis = data.frame(sapply(CFrancis, as.numeric))
CFrancis$Year = as.factor(CFrancis$Year)
CFrancis = missForest(CFrancis) #MISS_FOREST IMPUTATION
CFrancis <- CFrancis$ximp
CFrancis = cbind(CFrancis,"Winter"=rowMeans(CFrancis[c(6:8)]))
CFrancis = cbind(CFrancis,"Spring"=rowMeans(CFrancis[c(9:11)]))
CFrancis = cbind(CFrancis,"Summer"=rowMeans(CFrancis[c(12,13,2)]))
CFrancis = cbind(CFrancis,"Autumn"=rowMeans(CFrancis[c(3:5)]))
CFrancis = cbind(CFrancis,"Annual"=rowMeans(CFrancis[c(2:13)]))
write_xlsx(CFrancis,"D:\\Masters\\Provinces\\EC\\CFrancis-Prec.xlsx")
mk.test(CFrancis$Autumn)

CFrancis_radar = read_excel("D:\\Masters\\Provinces\\EC\\CFrancis-Prec.xlsx",sheet= 2)
radarchart(CFrancis_radar[2:18],seg = 3,title = "Cape Francis")

##EAST LONDON##
ELondon = read_excel("D:\\Masters\\Precipitation.xlsx",sheet= 12)
colnames(ELondon) = ELondon[2,]
ELondon = ELondon[-c(1:3),]
ELondon = data.frame(sapply(ELondon, as.numeric))
ELondon$Year = as.factor(ELondon$Year)
ELondon = missForest(ELondon) #MISS_FOREST IMPUTATION
ELondon <- ELondon$ximp
ELondon = cbind(ELondon,"Winter"=rowMeans(ELondon[c(6:8)]))
ELondon = cbind(ELondon,"Spring"=rowMeans(ELondon[c(9:11)]))
ELondon = cbind(ELondon,"Summer"=rowMeans(ELondon[c(12,13,2)]))
ELondon = cbind(ELondon,"Autumn"=rowMeans(ELondon[c(3:5)]))
ELondon = cbind(ELondon,"Annual"=rowMeans(ELondon[c(2:13)]))
write_xlsx(ELondon,"D:\\Masters\\Provinces\\EC\\ELondo-Prec.xlsx")


ELondon_radar = read_excel("D:\\Masters\\Provinces\\EC\\ELondo-Prec.xlsx",sheet= 2)
radarchart(ELondon_radar[2:18],seg = 3,title = "East London")

mk.test(ELondon$Autumn)

MFI = read_excel("D:\\Masters\\Provinces\\EC\\PCI.xlsx")
MFI = data.frame(t(MFI))
colnames(MFI) = MFI[1,]
MFI = MFI[-1,]
MFI = sapply(MFI,as.numeric)
boxplot(MFI,col="light blue",main="Annual PCI index by gauge station in the Eastern Cape province",ylab="PCI index")
abline(h=15, col="#CC3333")
