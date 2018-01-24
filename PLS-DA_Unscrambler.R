rm(list=ls())

#####
#This script has been redacted specifically for association of phenotypes to metabolites.
#Other uses should be plausible.

#This script mimics what The Unscrambler X do. The calculations and output plots expected.

#Instructions: the Excel should have a specific structure:
#In this file the Rows represent each Sample or Replicate.
#The First Column should have a name or identifier for every Sample or Replicate at the Column Name must be 'Name'
#Starting from the Second Column you should put all the information to be used as Predictors/Features, commonly the phenotypic traits.
#After the last Column with Predictors/Responses you should put all the information regarding Responses, commonly the metabolites measured.
#NOTE: the predictors should ALWAYS be Categorical Data, meaning groups, quantiles, yes/no, presence/absence, high/mid/low, etc.

#Data are stored in the selected output folder as sub folder representing every iteration process.
#The iterations remove data that falls outside the confidence ellipse calculated.
#The script re-do the calculations without this data, and filter again if needed.

#IGNORE errors related to ggplot (ej: removed x rows containing missing values (geom_xxxx))
#This are plotting errors and there are other plots who don't look pretty but shows the total data
#####


#Uncomment the line below to install the libraries 
#install.packages(c("sp","xlsx","pls","ggplot2","plotly","mvdalab","cowplot","tidyselect"))

#Loading Libraries
{
  library("sp")
  library("xlsx")
  library("pls")
  library("ggplot2")
  library("plotly")
  library("mvdalab")
  library("cowplot")
  library("tidyselect")
}


#Set working directory 
InputFolder="Folders/Input"
OutpuFolder="Folders/Output"


#Set name of the Excel File
NombreExcel="InputFile.xlsx"
#Set name of the Excel Sheet
ProjectName="ExcelSheet" #Nombre de la hoja a trabajar


setwd(paste0(InputFolder))

#Loading data into R. Dont change this.
Input=as.data.frame(read.xlsx(NombreExcel, header=TRUE, sheetName = ProjectName))
#Saving an original copy of the Input. Dont change this.
Original=Input

#Range of columns with Predictors/Features. Commonly started at column 2, since column 1 has the names or identifier.
ColumnaCategoria=c(2:7)
#Range of columns with numeric Predictors (not categorical data). Ignore this part as is purely for testing purposes.
ColumnaNumerica=c(NA)
#Column at which the metabolites or Responses variable starts. The script takes every column starting from this point as Response.
ColumnaMetabolitos=8

#You should stop changing the script from this point.


#Create Simple Ellipse (Hotellings)
simpleEllipse=function (x, y, alfa = 0.95, len = 200){
  N <- length(x)
  A <- 2
  mypi <- seq(0, 2 * pi, length = len)
  r1 <- sqrt(var(x) * qf(alfa, 2, N - 2) * (2 * (N^2 - 1)/(N * 
                                                             (N - 2))))
  r2 <- sqrt(var(y) * qf(alfa, 2, N - 2) * (2 * (N^2 - 1)/(N * 
                                                             (N - 2))))
  cbind(r1 * cos(mypi) + mean(x), r2 * sin(mypi) + mean(y))
}

#Blank Theme for ggplot
theme_blank <- theme(
  panel.grid.major = element_blank(),
  panel.grid.minor = element_blank(),
  panel.background = element_blank())


#    ____              _                  
#   |  _ \            | |                 
#   | |_) | ___   ___ | | ___  __ _ _ __  
#   |  _ < / _ \ / _ \| |/ _ \/ _` | '_ \ 
#   | |_) | (_) | (_) | |  __/ (_| | | | |
#   |____/ \___/ \___/|_|\___|\__,_|_| |_|
#                                         
outlayerLength=Inf
iterationNumber=0
while(!(outlayerLength==0)){

#Remover Filas correspondientes a Outlayers
if(exists("RemoveVar")){
#Cargar Datos nuevamente. Good practices.
print("Re reading Data")
Input=Original

iterationNumber=iterationNumber+1
print(paste("Iteration number",iterationNumber))
print(c("Removing",RemoveVar))
Input=Input[!(Input$Name %in% RemoveVar),]
Input$Name=factor(Input$Name)
}else{
RemoveVar=NULL
print(paste("Iteration number",iterationNumber))
print("Nothing to remove")}

#    ____  _     ____  
#   |  _ \| |   / ___| 
#   | |_) | |   \___ \ 
#   |  __/| |___ ___) |
#   |_|   |_____|____/ 
#                      
#Calculo Factores Correcci?n
Input[,(ColumnaMetabolitos:ncol(Input))]=apply((Input[,(ColumnaMetabolitos:ncol(Input))]), 2, function(x) x*(1/(sd(x)+0)))

#Generar matriz de datos
counterData=as.matrix(Input[,ColumnaMetabolitos:ncol(Input)])
rownames(counterData)=Input$Name

#Matriz Categorias
LogicalMatrix=as.data.frame(matrix(nrow=nrow(Input)))
for(x in ColumnaCategoria){
  for(i in unique(factor(Input[,x]))){
    LogicalMatrix[,i]=(ifelse(grepl(i , Input[,x]) , 1 , 0))
  }
}
LogicalMatrix$V1=NULL

#Fusion categor?as + predictor numerico
if(is.na(ColumnaNumerica[1])){
  print("There's no Numeric column to use as predictor")
}else{
  LogicalMatrix=cbind(LogicalMatrix,Input[,ColumnaNumerica])
}

for(i in 1:ncol(LogicalMatrix)){
  LogicalMatrix[,i]=LogicalMatrix[,i]*(1/(sd(LogicalMatrix[,i])+0))
}
LogicalMatrix[,1]
LogicalMatrix=as.matrix(LogicalMatrix)

plsrData <- plsr(LogicalMatrix ~ counterData, 
                 ncomp = 7, validation = "CV",
                 segments=nrow(Input), method = "oscorespls")

#Correlation Loadings for X & Y
cl=cor(model.matrix(plsrData),(scores(plsrData)[,1:2,drop=FALSE]))
clY=cor(LogicalMatrix, scores(plsrData)[,1:2, drop=FALSE])

#Inner and Outer Circle
radii=c(sqrt(1/2),1)

#Creating Dataframes
clplot=as.data.frame(cl)
clplotY=as.data.frame(clY)

#Computing Circles for ggplot
#Circle 1
theta <- seq(0, 2 * pi, length = 500)

xtemp1=radii[1]*cos(theta)
ytemp1=radii[1]*sin(theta)
circ1=as.data.frame(cbind(xtemp1,ytemp1))
colnames(circ1)=c("cx1","cy1")

#Circle 2
xtemp2=radii[2]*cos(theta)
ytemp2=radii[2]*sin(theta)
circ2=as.data.frame(cbind(xtemp2,ytemp2))
colnames(circ2)=c("cx2","cy2")

#Working
setwd(paste0(OutpuFolder))
#Crear Directorio
dir.create(paste(ProjectName,iterationNumber, sep="_"), showWarnings = FALSE)
setwd(paste(OutpuFolder, paste(ProjectName,iterationNumber, sep="_"),sep="/"))

#Raw Plot
p1=ggplot(data=clplot, aes(x=`Comp 1`, y=`Comp 2`)) + geom_point() +
  scale_x_continuous(limits=c(-1,1), breaks=seq(-1,1,by=0.1)) + 
  scale_y_continuous(limits=c(-1,1), breaks=seq(-1,1,by=0.1))
#Add Cross
p1=p1 + geom_hline(yintercept = 0, alpha=0.2) + geom_vline(xintercept = 0, alpha=0.2)
#Circle
p1=p1 + geom_polygon(data=circ1 ,aes(x=cx1, y=cy1),size=0.5,inherit.aes = F, colour="black", fill=NA, alpha=0.5) +
  geom_polygon(data=circ2, aes(x=cx2, y=cy2),size=0.5,inherit.aes = F, colour="black", fill=NA, alpha=0.5)
#Condition Points
p1=p1 + geom_point(data=clplotY, aes(x=`Comp 1`, y=`Comp 2`, color="red")) +
  geom_text(data=clplotY,nudge_x=0.02, nudge_y=0.02, fontface ="bold",aes(label=rownames(clplotY)))
#Labels
p1=p1 + labs(y =paste("Factor-2 (",substr(explvar(plsrData)[2],1,5),"%)", sep=""),
             x =paste("Factor-1 (",substr(explvar(plsrData)[1],1,5),"%)", sep=""))
#Other
p1= p1 + theme(legend.position="none")
#Chemicals Labels
p1=p1 + geom_text(data=clplot, nudge_y=0.05, size=3, aes(label=paste(1:nrow(clplot),substr(rownames(clplot),1,10), sep="-"))) +
  theme_blank

temp=ggplotly(p1)
htmlwidgets::saveWidget(as_widget(temp), paste(ProjectName,iterationNumber,"_Predicted-vs-Reference[PLS].html", sep=""))
save_plot(paste(ProjectName,iterationNumber,"_Predicted-vs-Reference[PLS].pdf", sep=""),p1, base_width=11, base_height=10)
write.xlsx(clplot, paste(ProjectName,iterationNumber,"_Predicted-vs-Reference[PLS].xlsx",sep=""),sheetName="clplotX")
write.xlsx(clplotY, paste(ProjectName,iterationNumber,"_Predicted-vs-Reference[PLS].xlsx",sep=""),sheetName="clplotY", append=TRUE)
write.xlsx(circ1, paste(ProjectName,iterationNumber,"_Predicted-vs-Reference[PLS].xlsx",sep=""),sheetName="InnerCircle", append=TRUE)
write.xlsx(circ2, paste(ProjectName,iterationNumber,"_Predicted-vs-Reference[PLS].xlsx",sep=""),sheetName="OuterCircle", append=TRUE)

#Raw Plot
p1v2=ggplot(data=clplot, aes(x=`Comp 1`, y=`Comp 2`)) +
  scale_x_continuous(limits=c(-1,1), breaks=seq(-1,1,by=0.1)) + 
  scale_y_continuous(limits=c(-1,1), breaks=seq(-1,1,by=0.1))
#Add Cross
p1v2=p1v2 + geom_hline(yintercept = 0, alpha=0.2) + geom_vline(xintercept = 0, alpha=0.2)
#Circle
p1v2=p1v2 + geom_polygon(data=circ1 ,aes(x=cx1, y=cy1),size=0.5,inherit.aes = F, colour="black", fill=NA, alpha=0.5) +
  geom_polygon(data=circ2, aes(x=cx2, y=cy2),size=0.5,inherit.aes = F, colour="black", fill=NA, alpha=0.5)
#Condition Points
p1v2=p1v2 + geom_point(data=clplotY, aes(x=`Comp 1`, y=`Comp 2`, color="red")) +
  geom_text(data=clplotY,nudge_x=0.02, nudge_y=0.02, fontface ="bold",aes(label=rownames(clplotY)))
#Labels
p1v2=p1v2 + labs(y =paste("Factor-2 (",substr(explvar(plsrData)[2],1,5),"%)", sep=""),
                 x =paste("Factor-1 (",substr(explvar(plsrData)[1],1,5),"%)", sep=""))
#Other
p1v2= p1v2 + theme(legend.position="none")
p1v2=p1v2 + geom_text(data=clplot, size=3, aes(label=seq(1,nrow(clplot),by=1))) +
  theme_blank
#Sin Chemical Labels
temp=ggplotly(p1v2)
htmlwidgets::saveWidget(as_widget(temp), paste(ProjectName,iterationNumber,"_Predicted-vs-Reference[PLS-NoLabel].html", sep=""))
save_plot(paste(ProjectName,iterationNumber,"_Predicted-vs-Reference[PLS-NoLabel].pdf", sep=""),p1v2, base_width=11, base_height=10)

#    ____   ____    _    
#   |  _ \ / ___|  / \   
#   | |_) | |     / _ \  
#   |  __/| |___ / ___ \ 
#   |_|    \____/_/   \_\
#                        
counterData=as.data.frame(counterData)

##PLOT PCA CON %. "Scores" Plot en Unscrambler.

#PCA. Eigenvalues and % of PC's equal to Unscrambler
pc1=suppressWarnings(pcaFit(counterData, scale=TRUE, ncomp=7))
labelPercent=paste(substr((pc1$Percents.Explained[1:2,1]),1,4),"%", sep="")

#Datos de Score y Loadings. Muy similar a Unscrambler
pc1Nipals=pca.nipals(counterData, ncomps=7, Iters=100)
dummy=data.frame(pc1Nipals$Scores)
pcaData=dummy[1:2]

#Score Plot SIN Hotellings T2-Test
p4=ggplot(data=pcaData,aes(x=as.numeric(pcaData$X1), y=as.numeric(pcaData$X2))) + 
  geom_point() + 
  geom_text(label=Input$Name, nudge_y=0.5) +
  labs(x=paste("PC-1 (",labelPercent[1],")", sep=""), y=paste("PC-2 (",labelPercent[2],")", sep=""))
#Add a Cross
p4=p4 + geom_hline(yintercept = 0, alpha=0.2) + geom_vline(xintercept = 0, alpha=0.2)

temp=ggplotly(p4)
htmlwidgets::saveWidget(as_widget(temp), paste(ProjectName,iterationNumber,"_[PCA-NoLimits].html", sep=""))
save_plot(paste(ProjectName,iterationNumber,"_[PCA-NoLimits].pdf", sep=""),p4, base_width=11, base_height=10)

#Add Limits
xlimits=abs(floor(min(pcaData$X1)))>abs(ceiling(max(pcaData$X1)))
xlimitsRes=ifelse(xlimits==TRUE,abs(floor(min(pcaData$X1))),abs(ceiling(max(pcaData$X1))))
ylimits=abs(floor(min(pcaData$X2)))>abs(ceiling(max(pcaData$X2)))
ylimitsRes=ifelse(ylimits==TRUE,abs(floor(min(pcaData$X2))),abs(ceiling(max(pcaData$X2))))
p4= p4 + scale_x_continuous(limits=c(-1*xlimitsRes,xlimitsRes), breaks=seq(-1*xlimitsRes,xlimitsRes, by=1)) +
  scale_y_continuous(limits=c(-1*ylimitsRes,ylimitsRes), breaks=seq(-1*ylimitsRes,ylimitsRes, by=1)) +
  theme_blank

temp=ggplotly(p4)
htmlwidgets::saveWidget(as_widget(temp), paste(ProjectName,iterationNumber,"_Scores[PCA-NoHotellings].html", sep=""))
save_plot(paste(ProjectName,iterationNumber,"_Scores[PCA-NoHotellings].pdf", sep=""),p4, base_width=11, base_height=10)
write.xlsx(pcaData, paste(ProjectName,iterationNumber,"_Scores[PCA-NoHotellings].xlsx",sep=""),sheetName="PCA-Nipals")

#Score plot with Hotelling T2-Test
#Generar datos de circulos para Hotelling T2-Test
ellTest=as.data.frame(simpleEllipse(pcaData[,1],pcaData[,2]))

#Gr?fico de Score CON Hotelling T2-Test
p5=ggplot(data=pcaData, aes(x=X1, y=X2)) + geom_point() + 
  geom_polygon(data=ellTest,aes(x=V1, y=V2),inherit.aes = F, colour="black", fill=NA, alpha=0.5) +
  geom_hline(yintercept = 0) + geom_vline(xintercept = 0) +
  labs(x=paste("PC-1 (",labelPercent[1],")", sep=""), y=paste("PC-2 (",labelPercent[2],")", sep="")) +
  geom_text(label=Input$Name, nudge_y=0.5) +
  theme_blank

temp=ggplotly(p5)
htmlwidgets::saveWidget(as_widget(temp), paste(ProjectName,"_Scores[PCA-Hotellings-NoLimits].html", sep=""))
save_plot(paste(ProjectName,"_Scores[PCA-Hotellings-NoLimits].pdf", sep=""),p5, base_width=11, base_height=10)
write.xlsx(ellTest, paste(ProjectName,iterationNumber,"_Scores[PCA-Hotellings].xlsx",sep=""),sheetName="PCA-Nipals-Hotellings")

#Computing Limits
xlimits=abs(floor(min(ellTest$V1)))>abs(ceiling(max(ellTest$V1)))
xlimitsRes=ifelse(xlimits==TRUE,abs(floor(min(ellTest$V1))),abs(ceiling(max(ellTest$V1))))
ylimits=abs(floor(min(ellTest$V2)))>abs(ceiling(max(ellTest$V2)))
ylimitsRes=ifelse(ylimits==TRUE,abs(floor(min(ellTest$V2))),abs(ceiling(max(ellTest$V2))))
p5= p5 + scale_x_continuous(limits=c(-1*xlimitsRes,xlimitsRes), breaks=seq(-1*xlimitsRes,xlimitsRes, by=1)) +
  scale_y_continuous(limits=c(-1*ylimitsRes,ylimitsRes), breaks=seq(-1*ylimitsRes,ylimitsRes, by=1)) +
  theme_blank

temp=ggplotly(p5)
htmlwidgets::saveWidget(as_widget(temp), paste(ProjectName,iterationNumber,"_Scores[PCA-Hotellings].html", sep=""))
save_plot(paste(ProjectName,iterationNumber,"_Scores[PCA-Hotellings].pdf", sep=""),p5, base_width=11, base_height=10)

######
clPCA=as.data.frame(cor(model.matrix(plsrData),pcaData))
#Raw Plot
p7=ggplot(data=clPCA, aes(x=X1, y=X2)) + geom_point() +
  scale_x_continuous(limits=c(-1,1), breaks=seq(-1,1,by=0.1)) + 
  scale_y_continuous(limits=c(-1,1), breaks=seq(-1,1,by=0.1))
#Add Cross
p7=p7 + geom_hline(yintercept = 0, alpha=0.2) + geom_vline(xintercept = 0, alpha=0.2)
#Circle
p7=p7 + geom_polygon(data=circ1 ,aes(x=cx1, y=cy1),size=0.5,inherit.aes = F, colour="black", fill=NA, alpha=0.5) +
  geom_polygon(data=circ2, aes(x=cx2, y=cy2),size=0.5,inherit.aes = F, colour="black", fill=NA, alpha=0.5)
#Labels
p7=p7 + labs(y =paste("Factor-2 (",substr(explvar(plsrData)[2],1,5),"%)", sep=""),
             x =paste("Factor-1 (",substr(explvar(plsrData)[1],1,5),"%)", sep=""))
#Other
p7= p7 + theme(legend.position="none")
#Chemicals Labels
p7=p7 + geom_text(data=clPCA, nudge_y=0.05, size=3, aes(label=paste(1:nrow(clPCA),substr(rownames(clPCA),1,10), sep="-"))) +
  theme_blank

temp=ggplotly(p7)
htmlwidgets::saveWidget(as_widget(temp), paste(ProjectName,iterationNumber,"_CorrelationLoading[PCA].html", sep=""))
save_plot(paste(ProjectName,iterationNumber,"_CorrelationLoading[PCA].pdf", sep=""),p7, base_width=11, base_height=10)
write.xlsx(clPCA, paste(ProjectName,iterationNumber,"_CorrelationLoading[PCA].xlsx",sep=""),sheetName="CorrelationLoadingPCA")

#Raw Plot
p7v2=ggplot(data=clPCA, aes(x=X1, y=X2)) + 
  scale_x_continuous(limits=c(-1,1), breaks=seq(-1,1,by=0.1)) + 
  scale_y_continuous(limits=c(-1,1), breaks=seq(-1,1,by=0.1))
#Add Cross
p7v2=p7v2 + geom_hline(yintercept = 0, alpha=0.2) + geom_vline(xintercept = 0, alpha=0.2)
#Circle
p7v2=p7v2 + geom_polygon(data=circ1 ,aes(x=cx1, y=cy1),size=0.5,inherit.aes = F, colour="black", fill=NA, alpha=0.5) +
  geom_polygon(data=circ2, aes(x=cx2, y=cy2),size=0.5,inherit.aes = F, colour="black", fill=NA, alpha=0.5)
#Labels
p7v2=p7v2 + labs(y =paste("Factor-2 (",substr(explvar(plsrData)[2],1,5),"%)", sep=""),
                 x =paste("Factor-1 (",substr(explvar(plsrData)[1],1,5),"%)", sep=""))
#Other
p7v2= p7v2 + theme(legend.position="none")
p7v2=p7v2 + geom_text(data=clPCA, size=3, aes(label=seq(1,nrow(clPCA),by=1))) +
  theme_blank
#No Chemical Label
temp=ggplotly(p7v2)
htmlwidgets::saveWidget(as_widget(temp), paste(ProjectName,iterationNumber,"_CorrelationLoading[PCA-NoLabel].html", sep=""))
save_plot(paste(ProjectName,iterationNumber,"_CorrelationLoading[PCA-NoLabel].pdf", sep=""),p7v2, base_width=11, base_height=10)

##Vip 
VIP=as.data.frame(loading.weights(plsrData)[,1:2, drop=FALSE])
VIP$Name=paste(1:nrow(VIP),substr(rownames(VIP),1,10), sep="-")
VIP$Compound=paste(1:nrow(VIP),rownames(VIP), sep="-")
VIP$distance=sqrt(((VIP$`Comp 1`)^2)+((VIP$`Comp 2`)^2))
VIP=VIP[order(-VIP$distance),]
#Plot VIP Ordenados
p6=ggplot(data=VIP, aes(x=reorder(VIP$Name, -distance), y=distance)) + 
  geom_col(width=0.5) + theme(axis.text.x=element_text(angle=90))
#Labels
p6= p6 + labs(y = "Distancia", x = "Compuestos")
#Limites
p6= p6 + scale_y_continuous(expand=c(0,0),limits=c(0,max(VIP$distance)),
                            breaks=seq(0,round(max(VIP$distance), digits=1), by=0.05))
temp=ggplotly(p6)
htmlwidgets::saveWidget(as_widget(temp), paste(ProjectName,"_[VIP].html", sep=""))
save_plot(paste(ProjectName,iterationNumber,"_[VIP].pdf", sep=""),p6, base_width=14, base_height=5)

#Guardar Datos de VIP
VIP$Name=NULL
write.xlsx(VIP, paste(ProjectName,iterationNumber,"_[VIP].xlsx",sep=""),row.names=FALSE)

setwd(paste0(InputFolder))

#Setting Outlayers
Outlayers=pcaData[point.in.polygon(pcaData$X1, pcaData$X2, ellTest$V1, ellTest$V2) == 0,]
assign("RemoveVar",c(RemoveVar,rownames(Outlayers)))
outlayerLength=length(row.names(Outlayers))
}
##FIN##

