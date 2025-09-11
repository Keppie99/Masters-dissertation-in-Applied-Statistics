library(readxl)
library(missForest)
library(trend)
library(precintcon)
library(writexl)
library(fmsb)
library(missForest)

#####PRECIPITATION#####
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