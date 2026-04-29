rm(list=ls())

#Load the required libraries

library(readxl)
library(xlsx)
library(fUnitRoots)
library(tseries)
library(forecast)
library(sarima)
library(lmtest)
library(rsq)
library(polynom)
library(fracdiff)
library(hypergeo)
library(sandwich)
library(longmemo)
library(abind)
library(fractaldim)
library(FractalParameterEstimation)
library(liftLRD)
library(pracma)
library(arfima)
library(afmtools)
library(rugarch)
library(stats)
library(LongMemoryTS)
# set working directory in which you want to open and save your extractions

setwd("C:/Users/...........................")

#######################################
# Reading (opening) the data file and choosing specific data from it

#Excel sheet contaiting all the données: File
File<- read_excel("C:/Users/,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,.xlsx")
#Import from the File the river flow datasetand give it a name. In this code it was identified as c02 
c02 <- read_excel("C:/Users/,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,.xlsx",range = cell_cols("B") )
#Import from thje File the time span
Time <- read_excel("C:/Users/,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,.xlsx",range = cell_cols("A"))

#######################################
#Analyzing the data file

sf=summary(File)# provides a minimum value, maximum value, median and the mean for all the embedded data
sf
class(c02)# class of the data embedded in the file should be time series
c02.ts= ts(c02)# creating time series for the variable in the data file- this is useful for strudying the data in the file not taking into consideration the value of the year instead of this it takes them as step points

class(c02.ts)# provides the classes in each column detected in the data file

#######################################
# Define 

C02.month=ts(c02,start=c(1883,1),end = c(2018,1),freq=12)#Defining a time series data inserting their time value and step #This uses all the data. You can select a smaller number by specifying a nearlier end date using the parameter end
class(C02.month)
sd=summary(C02.month)
sd
plot(C02.month)# plotting the time series 
bp=boxplot(C02.month ~ cycle(C02.month),xlab="Months",pars = list(boxwex = 0.8, staplewex = 0.5, outwex = 0.5),col = "white",ylab = "m3/sec")#You can see  the seasonal effects in the boxplot

bpstats=bp$stats#	a matrix, each column contains the extreme of the lower whisker, the lower hinge, the median, the upper hinge and the extreme of the upper 
#whisker for one group/plot. If all the inputs have the same class attribute, so will this component.
bpconf=bp$conf#a matrix where each column contains the lower and upper extremes of the notch.

#####seasonal effect conclusion:
##increase in mars et avril and decrease in aout et septembre 

#######################################
# Visualize the seasonal effect detected in specefic months

C02.janv=window(C02.month, start=c(1883,1),frequency=TRUE)
C02.janv
C02.fev=window(C02.month, start=c(1883,2),frequency=TRUE)
C02.fev
C02.mars=window(C02.month, start=c(1883,3),frequency=TRUE)
C02.mars
C02.avril=window(C02.month, start=c(1883,4),frequency=TRUE)
C02.avril
C02.mai=window(C02.month, start=c(1883,5),frequency=TRUE)
C02.mai
C02.june=window(C02.month, start=c(1883,6),frequency=TRUE)
C02.june
C02.juillet=window(C02.month, start=c(1883,7),frequency=TRUE)
C02.juillet
C02.aout=window(C02.month, start=c(1883,8),frequency=TRUE)
C02.aout
C02.septembre=window(C02.month, start=c(1883,9),frequency=TRUE)
C02.septembre
C02.octobre=window(C02.month, start=c(1883,10),frequency=TRUE)
C02.octobre
C02.novembre=window(C02.month, start=c(1883,11),frequency=TRUE)
C02.novembre
C02.decembre=window(C02.month, start=c(1883,12),frequency=TRUE)
C02.decembre

plot(C02.janv)
plot(C02.fev)
plot(C02.mars)
plot(C02.avril)
plot(C02.mai)
plot(C02.june)
plot(C02.juillet)
plot(C02.aout)
plot(C02.septembre)
plot(C02.octobre)
plot(C02.novembre)
plot(C02.decembre)

#######################################
#Aggregation
#used to detect the total number of C02(river flow) in each year and give a view without the seasonal effect

ag=aggregate(C02.month);ag# summation of flow in each year
plot(aggregate(C02.month))# used To get a clearer view of the trend, the seasonal effect can be removed by aggregating the data




#######################################

#Mean annual rate

C02.AR=aggregate(C02.month)/12# average in each year=mean annual rate
C02.ARcheck=aggregate(C02.month,FUN = mean)# =same of C02.AR used to check the frequency
C02.AR
C02.ARcheck
plot(C02.AR)
aggregatecheck=C02.AR-C02.ARcheck;aggregatecheck# zero values
plot(aggregatecheck)


##### Mean annual rate conclusion:
## an increase in the mean annual rate and a decrease is detected in 1987  

#######################################

# Month attribution:
#This describe the increase or the decrease of each month's mean in the terms of percentage over the mean of all months in the covered period

janv.ratio=mean(C02.janv)/mean(C02.month)
fev.ratio=mean(C02.fev)/mean(C02.month)
mars.ratio=mean(C02.mars)/mean(C02.month)
avril.ratio=mean(C02.avril)/mean(C02.month)
mai.ratio=mean(C02.mai)/mean(C02.month)
june.ratio=mean(C02.june)/mean(C02.month)
juillet.ratio=mean(C02.juillet)/mean(C02.month)
aout.ratio=mean(C02.aout)/mean(C02.month)
septembre.ratio=mean(C02.septembre)/mean(C02.month)
octobre.ratio=mean(C02.octobre)/mean(C02.month)
novembre.ratio=mean(C02.novembre)/mean(C02.month)
decembre.ratio=mean(C02.decembre)/mean(C02.month)

months.ratio=rbind(janv.ratio,fev.ratio,mars.ratio,avril.ratio,mai.ratio,june.ratio,juillet.ratio,aout.ratio,septembre.ratio,octobre.ratio,novembre.ratio,decembre.ratio)
months.ratio
plot(months.ratio)

##############
#This describe the increase or the decrease of each month's variance in the terms of percentage over the variance of all months in the covered period

janv.ratio=var(C02.janv)/var(C02.month)
fev.ratio=var(C02.fev)/var(C02.month)
mars.ratio=var(C02.mars)/var(C02.month)
avril.ratio=var(C02.avril)/var(C02.month)
mai.ratio=var(C02.mai)/var(C02.month)
june.ratio=var(C02.june)/var(C02.month)
juillet.ratio=var(C02.juillet)/var(C02.month)
aout.ratio=var(C02.aout)/var(C02.month)
septembre.ratio=var(C02.septembre)/var(C02.month)
octobre.ratio=var(C02.octobre)/var(C02.month)
novembre.ratio=var(C02.novembre)/var(C02.month)
decembre.ratio=var(C02.decembre)/var(C02.month)

months.ratiovar=rbind(janv.ratio,fev.ratio,mars.ratio,avril.ratio,mai.ratio,june.ratio,juillet.ratio,aout.ratio,septembre.ratio,octobre.ratio,novembre.ratio,decembre.ratio)
months.ratiovar
plot(months.ratiovar)

#######################################

#Decomposition
#Decomposition model:
#additive:Xt=mt+st+zt 
#multiplicative:Xt=mt.st+zt : used when the seasonal effect increases with the increase of the trend
# script line 89 scripted under :# Visualize the seasonal effect detected in specefic months: check script line 75 till 86 to make the decision
#Xt:observed data, mt:trend, st:seasonal, zt:random variable with mean equals to zero

Decomposition=decompose(C02.month,type = c("multiplicative"),filter = NULL);Decomposition
plot(Decomposition)
Decomposition$x #The original series
Decomposition$seasonal #The seasonal component which is the repeated seasonal figure. using a moving average method
Decomposition$trend# The trend componenet. using a moving average method
Decomposition$random# The remainder part which is an estimate of the random process is obtained from the original time series using estimates of the trend and seasonal effects. it is a residual error series and it is not precisely a realisation of the random process zt  but rather an estimate of that realisation.
Decomposition$figure # the seasonal figure which means seasonal pattern = months.ratio

plot(Decomposition$x)
plot(Decomposition$seasonal)
plot(Decomposition$figure)
plot(Decomposition$trend)
plot(Decomposition$random)

#Checks on random part
decompositionmean=mean(Decomposition$random[7:1615]);decompositionmean#check the value: distribution is around it depending on the variance
decompositionvar=var(Decomposition$random[7:1615]);decompositionvar#check the value, smaill variance must be detected 

#Effect of seasonality on trend
ts.plot(cbind(Decomposition$trend,Decomposition$trend*Decomposition$seasonal),lty=1:2)# dotted is the seasonal effect on trend

#Evaluation of the seasonal pattern accuracy
#decides if the type of the decomposition is accepted by calculating the errors 
error.Decomposition=Decomposition$figure-months.ratio
error.Decomposition
plot(error.Decomposition)

#Autocorrelation usage in detecting the:

#type of time analysis: 
#shot:Time series characterised by sharp decaying ACF will usually exhibit a sharp pattern and can be classified as short-memory time or long:Time series characterised by slowly decaying ACF will usually exhibit a smooth pattern and can be classified as long-memory time
#trend effect: 
#high value at lag 1 & decrease with the increase of lags
#seasonality:autocorrelation will be larger with peaks at seasonal data

acf(C02.month)#check type of time series analysis if it is short or long
acfdatabase=acf(C02.month, type= c("correlation"), lag.max = 48,plot = T);acfdatabase
show(acfdatabase)
attributes(acfdatabase)
acfdatabase$acf#acfvalues
acfdatabase$lag#lagvalues
acfdatabase.df=as.data.frame(acfdatabase$acf,acfdatabase$lag)
acfdatabase.df#includes acf and lags values of the database

# Check if we need AR model or we can use other features;
damagedcosine=acf(Decomposition$random[7:1615],lag.max = 48)#damaged cosine shape is a characteristic of AR model
damagedcosine.df=as.data.frame(damagedcosine$acf,damagedcosine$lag);damagedcosine.df
pacfdatabase=pacf(C02.month,lag.max = 30)
pacfdatabase
pacfdatabase.df=as.data.frame(pacfdatabase$acf,pacfdatabase$lag);pacfdatabase.df

##### Decomposition & Autocorrelation usage conclusion:
## a multiplicative model is applied and accepted due to the script line 166 :evaluation of the seasonal pattern accuracy
##short time analysis
##there is low trend effect
## there is an effect for seasonality
##AR model is required in the study 

#######################################

#Differencing:

#Determine if differencing is required or not by Determining the stationarity of the time series 

#1-Unit Root Test Interface
#0: unit root test for stationarity
#provides acf,pacf and regression test tau residuals
#urkpssTest(x, type = c("mu", "tau"), lags = c("short", "long", "nil"), use.lag = NULL, doplot = TRUE)

urkpsstestresult=urkpssTest(C02.month, type = c("tau"), lags = c("short"),use.lag = NULL, doplot = TRUE)
attributes(urkpsstestresult)
urkpsstestresult
urkpsstestcriticalvalue=urkpsstestresult@test[["test"]]@cval;urkpsstestcriticalvalue
urkpssteststatistical=urkpsstestresult@test[["test"]]@teststat;urkpssteststatistical
urkpsstestresiduals=urkpsstestresult@test[["test"]]@res
ts.plot(urkpsstestresiduals)
urkpsstestdecision=if(urkpssteststatistical<0.216){
  print("null hypothesis is accepted; the series is stationary.")
} else {
  print("null hypothesis is rejected; the series is non-stationary. ")
}
urkpsstestdecision
urkpsstestresidualsacf=acf(urkpsstestresiduals)
urkpsstestresidualsacf$acf
urkpsstestresidualsacf$lag
acfurkpsstestresiduals.df=as.data.frame(urkpsstestresidualsacf$acf,urkpsstestresidualsacf$lag);acfurkpsstestresiduals.df
urkpsstestresidualspacf=pacf(urkpsstestresiduals);urkpsstestresidualspacf
urkpsstestresidualspacf$acf
urkpsstestresidualspacf$lag
pacfurkpsstestresiduals.df=as.data.frame(urkpsstestresidualspacf$acf,urkpsstestresidualspacf$lag);pacfurkpsstestresiduals.df

#Kpss test indicates the time series stationarity status and indicates with the value of pvalue if differencing is requried
#kpss.test(x, null = c("Level", "Trend"), lshort = TRUE)
#if the Pvalue<0.05 and statistic level  > critical  then the null hypothesis is rejected; the series is non-stationary. 
#if the Pvalue>0.05 and statistic level  < critical  then the null hypothesis is accepted; the series is stationary. 

#critical value of level
#Sample 	1%	  intercept	    linear trend	  
#alpha = 0.1	  0.347	  	       0.119	
#alpha = 0.05   0.463		         0.146	
#alpha = 0.01	  0.739	         	 0.216	
	
#Kpss Trend test

kpsstesttrend=kpss.test(C02.month, null = c("Trend"),lshort = TRUE)
kpsstesttrend$statistic# value compared with the critical
kpsstesttrend$parameter
kpsstesttrend$p.value # value compared with 0.05

KpssTesttrendpvalue=c(kpsstesttrend$p.value)
KpssTesttrendpvalue
kpssTesttrendstatistic=c(kpsstesttrend$statistic)
kpssTesttrendstatistic
pCriteria=c(0.05)
pCriteria
kpsstest.resulttrend=ifelse(KpssTesttrendpvalue<pCriteria && kpssTesttrendstatistic>0.739,"Null hypothesis is rejected for trend and Time series is non stationary","Trend:Time series is stationary")
kpsstest.resulttrend


#Kpss Level test

kpsstestLevel=kpss.test(C02.month, null = c("Level"),lshort = TRUE)
kpsstestLevel$statistic# value compared with the critical
kpsstestLevel$parameter
kpsstestLevel$p.value # value compared with 0.05

KpssTestLevelpvalue=c(kpsstestLevel$p.value)
KpssTestLevelpvalue
kpssTestLevelstatistic=c(kpsstestLevel$statistic)
kpssTestLevelstatistic
pCriteria=c(0.05)
pCriteria
kpsstest.resultLevel=ifelse(KpssTestLevelpvalue<pCriteria && kpssTestLevelstatistic>0.739,"Null hypothesis is rejected for level and Time series is non stationary","Level: Time series is stationary")
kpsstest.resultLevel

#3-ADf test

#Critical values for Dickey-Fuller t-distribution.
#        Without trend  	With trend
#Sample size	1%	5%	    1%	  5%
#T = 25	  -3.75	-3.00	  -4.38	  -3.60
#T = 50	  -3.58	-2.93	  -4.15	  -3.50
#T = 100	-3.51	-2.89	  -4.04	  -3.45
#T = 250	-3.46	-2.88	  -3.99	  -3.43
#T = 500	-3.44	-2.87	  -3.98	  -3.42
#T =+500  -3.43	-2.86	  -3.96	  -3.41
#if the Pvalue >0.05 and adf statistic > critical then it is non stationatry time series
#if the pvalue <0.05 and adf statistic < critical then it is  stationatry time series

#adf.test(x, alternative = c("stationary", "explosive"),
#         k = trunc((length(x)-1)^(1/3)))

# Consequtive adf test time serious
adfcons=adf.test(C02.month,alternative = c("stationary"), k = trunc((length(C02.month)-1)^(1/3)))#k=1 can be used too 
adfcons
adfcons$statistic
adfcons$p.value
# Null hypothenous as pvalue <0.05 and adf statistic < critical  then it is  stationatry time series

adfpvalue=c(adfcons$p.value)
adfpvalue
adfstatistic=c(adfcons$statistic)
adfstatistic
adfcritical= -3.43
Criteria=c(0.05)
Criteria
adfresult=ifelse(adfpvalue>Criteria &&adfstatistic > adfcritical,"Null hypothesis is accepted for not time trend and there is a unit root and Time series is non stationary","Time series for not time trend test is stationary")
adfresult

#Seasonal adf test time series
adftestseas=adfTest(C02.month,lags=12,type=c("ct"),title = NULL, description = NULL)# seasonal effect
adftestseas
#the Pvalue>0.05 and adf statistic > critical then it is non stationatry seasonal time series

adfseaspvalue=c(adftestseas@test$p.value)
adfseaspvalue
adfseasstatistic=c(adftestseas@test$statistic)
adfseasstatistic
pCriteria=c(0.05)
pCriteria
adfseascritical= -3.96
adfseasresult=ifelse(adfseaspvalue>pCriteria &&adfseasstatistic > adfseascritical,"Null hypothesis is accepted for intercept and a time trend and there is a unit root and Time series is non stationary","Time series for intercept and a time trend  is stationary")
adfseasresult

#Evaluating the integration value 

#determine the number of differences required to make the time series stationary removing seasonality
#estimates the number of seasonal differences necessary.

integrationseasonalvalue=nsdiffs(C02.month,alpha=0.05, test = c("seas"),max.D = 2);integrationseasonalvalue

#Stationary with seasonality:
stationarywithseasonality=diff(C02.month,differences= integrationseasonalvalue);stationarywithseasonality
plot(stationarywithseasonality)

#Stationary removing seasonality:
C02.seasonaladjusted=C02.month-Decomposition$seasonal
stationaryremovingseasonality=diff(C02.seasonaladjusted,differences=12)
plot(stationaryremovingseasonality)










###type of time analysis: 

#shot:Time series characterised by sharp decaying ACF will usually exhibit
#a sharp pattern and can be classified as short-memory time or 
#long:Time series characterised by slowly decaying ACF will usually exhibit
#a smooth pattern and can be classified as long-memory time

#In the time domain, a hyperbolically decaying autocovariance function 
#characterizes the presence of long memory
acf(C02.month)$acf[2]#check type of time series analysis if it is short or long
acf(C02.month, type= c("correlation"), lag.max = 30)$acf[2]# used to detect seasonality for 2 years

#ARFIMA Spectral Analysis:
#used to identify the spectral analysis of the time series to identify the
#period of the time depending on its frequency 
#Simply estimate the periodogram via the Fast Fourier Transform.

#In the frequency domain, the long memory is present when the spectral 
#density function approaches infinity at low frequencies.Suppose f(X) is
#the spectral density function. The series xt is said to exhibit long memory
pa=per.arfima(C02.month)# used with hurst exponent to determine LRM
pa
periodogramspectrum=pa$spec
periodogramfrequency=pa$freq
plot(pa$freq, pa$spec, type="l", xlab="frequencies", 
     ylab="spectrum", main="Periodogram")
max(pa$spec)#choose the frequency that corresponds to this value 
# then calculate the period as 1/ frequency it takes number of time periods for a complete cycle. 
plot(Peri(C02.month))
period1=1/0.51
period2=1/1.1   
period1 
period2

#spectrum analysis
#The spectrum function estimates the spectral density of a time series.
#spectrum(x, ..., method = c("pgram", "ar"))

specdata=spectrum(C02.month)#"Raw Periodogram" spectrum by default
spectrumrawperiodogramfrequency=specdata$freq
spectrumrawperiodogramspectrum=specdata$spec
spectrumrawperiodogramdegreeoffreedom=specdata$df#The distribution of the spectral density estimate can be approximated by a (scaled) chi square distribution with df degrees of freedom
spectrumrawperiodogrambandwidth=specdata$bandwidth#The equivalent bandwidth of the kernel smoother 
specdata$method
spectrumrawperiodogramtaper=specdata$taper
specdata$pad
specdata$detrend
specdata$demean

specdataar=spectrum(C02.month,method = "ar")# ar method#Fits an AR model to x (or uses the existing fit) and computes (and by default plots) the spectral density of the fitted model.
spectrumARfrequency=specdataar$freq
spectrumARspectrum=specdataar$spec
specdataar$method

plot(specdataar$freq,specdataar$spec) # ar method
plot(specdata$freq,specdata$spec) # pgram method

#########################################################################

###estimation of fractioned integration 

##estimation of fractioned integration using periodgram
#fds=fdSperio(C02.month)
#fds$d
#sperio.d=fds$d;sperio.d#Sperio estimate ( d estimated)
#sperio.asymptotic.sd=fds$sd.as;sperio.asymptotic.sd#asymptotic standard deviation
#sperio.deviation.sd=fds$sd.reg;sperio.deviation.sd# standard error deviation

##GPH estimator of fractional difference parameter d.
#gph log-periodogram estimator of Geweke and Porter-Hudak (1983) (GPH) and Robinson (1995a) for memory parameter d.
# where 0<delta<1.
gph(C02.month, m=floor( 1+length(C02.month)^0.8), l = 1)
gph=0.2568565
##Multivariate local Whittle estimation of long memory parameters.
#GSE Estimates the memory parameter of a vector valued long memory process.
# where 0<delta<1.
GSE(C02.month, m=floor( 1+length(C02.month)^0.8), l = 1)#GSE Estimates the memory parameter of a vector valued long memory process.
GSE=0.3210783 
 
##Modified local Whittle estimator of fractional difference parameter d.
#Hou.Perron Modified semiparametric local Whittle estimator of Hou and Perron (2014). Estimates memory parameter robust to low frequency contaminations.
# where 0<delta<1.
HP=Hou.Perron(C02.month, m=floor( 1+length(C02.month)^0.8))
HP$d
HP.d=HP$d;HP.d#???fractional integration value
HP.se=HP$s.e.;HP.se# standard error
##Local Whittle estimator of fractional difference parameter d.
LWnone=local.W(C02.month, m=floor( 1+length(C02.month)^0.8), int = c(-0.5, 2.5), taper = c("none"), diff_param = 1, l = 1)#standard local Whittle estimator of Robinson (1995)
LWnone$d
LWnone.d=LWnone$d;LWnone.d#???fractional integration value
LWnone.se=LWnone$s.e.;LWnone.se
LWvelasco=local.W(C02.month, m=floor( 1+length(C02.month)^0.8), int = c(-0.5, 2.5), taper = c("Velasco"), diff_param = 1, l = 1)#he tapered version of Velasco (1999),
LWvelasco$d
LWvelasco.d=LWvelasco$d;LWvelasco.d
LWvelasco.se=LWvelasco$s.e.;LWvelasco.se
LWHC=local.W(C02.month, m=floor( 1+length(C02.month)^0.8), int = c(-0.5, 2.5), taper = c("HC"), diff_param = 1, l = 1)#the differenced and tapered estimator of Hurvich and Chen (2000)
LWHC$d
LWHC.d=LWHC$d;LWHC.d
LWHC.se=LWHC$s.e.;LWHC.se
##Local polynomial Whittle plus noise estimator
#LPWN calculates the local polynomial Whittle plus noise estimator of Frederiksen et al. (2012).
LPWN=LPWN(C02.month, m=floor( 1+length(C02.month)^0.65), R_short = 4, R_noise = 0)
LPWN
LPWNd=LPWN[["parameters"]][["d"]]
LPWNd
LPWNs.e=LPWN[["standard_errors"]][["d"]]
LPWNs.e

## Maximum and minimum value of the integration fractional d value
dmaxwithvelsaco=max(sperio.d,gph,GSE,HP.d,LWnone.d,LWvelasco.d,LWHC.d,LPWNd)
dmaxwithvelsaco
dmaxwithoutvelsacoandHC=max(gph,GSE,HP$d,LWnone$d,LPWNd)
dmaxwithoutvelsacoandHC

dminwithvelsaco=min(gph,GSE,HP$d,LWnone$d,LWvelasco$d,LWHC$d,LPWNd)
dminwithvelsaco
dminwithoutvelsacoandHC=min(gph,GSE,HP$d,LWnone$d,LPWNd)
dminwithoutvelsacoandHC
#Range of d 
dwithvelsacoandHC=c(dminwithvelsaco,dmaxwithvelsaco)
dwithvelsacoandHC
dwithoutvelsacoandHC=c(dminwithoutvelsacoandHC,dmaxwithoutvelsacoandHC)
dwithoutvelsacoandHC
###################################################################

###The autocorrelation of the time series using one of the above d estimation values
# we can use one of the 6 above tests value in this case to indicate stationarity of a time series
#Fast fractional differencing procedure of Jensen and Nielsen 

acfusingfractionalintegration=acf(fdiff(C02.month,d=2))#Auto correlation using fractioned integrated value of the periodgram and varies with the standard error
acfusingfractionalintegration$acf
acfusingfractionalintegration$lag
acfusingfractionalintegration.df=as.data.frame(acfusingfractionalintegration$acf,acfusingfractionalintegration$lag)
acfusingfractionalintegration.df

###Hurst exponent using R/S analysis
#The Hurst exponent (H) is a measure of long memory of time series [23]. 
#It relates to the autocorrelations of the time series and the rate at which 
#these values decrease as the lag increases. The relationship between df and H is:
#df=H???0.5; 
#if H>0.5, it would indicate a long-memory time series; if H<0.5, it can be considered as an 
#intermediate-memory time series. When H=0.5, it would indicate a random walk
# the hurst exponential >0.5 then it is LRM The accuracy of the forecasts also increases 
#when the Hurst exponent is higher than 0.5

H=hurstexp(C02.month)#Calculates the Hurst exponent using R/S analysis. 
hurstexponent=as.data.frame(H)
hurstexponent#includes all hurst exponent estimations 
H$Hs#Simple R/S Hurst estimation
H$Hrs#corrected R over S Hurst exponent
H$He#Empirical Hurst exponent
H$Hal#Corrected empirical Hurst exponent
H$Ht#Theoretical Hurst exponent

###estimations of fractional d value:
##For SARFIMA a stationary process, d varies between ???0.5 and 0.5, with d=0 indicating short memory,
#???0.5<d<0 indicating intermediate memory, and 0<d<0.5 indicating long memory
##For ARFIMA (p, d *, q), where d * = d + df. Most commonly, 
#df??? (???0.5, 0.5) is the fractional part, and d??? 0 always is the integer part. #df=H???0.5; 
#if H>0.5, it would indicate a long-memory time series; if H<0.5, it can be considered as 
#an intermediate-memory time series. When H=0.5, it would indicate a random walk

#d values that are calculated from the 6 tests
dwithvelsacoandHC
dwithoutvelsacoandHC
dARFIMA=c(0.01,0.5)# d value from previous 6 tests added with the standard error
df=H$Hs-0.5#calculated bysimple R/S hurst
df
dfcorrected=H$Hrs-0.5#calculated by corrected R/S hurst
dfcorrected
dstarARFIMA=dARFIMA+df#calculated bysimple R/S hurst
dstarARFIMA
dstarARFIMAcorrected=dARFIMA+dfcorrected
dstarARFIMAcorrected

#######################################################################
###Qu test for true long memory against spurious long memory.
#Qu.test Test statistic of Qu (2011) for the null hypotesis of true long memory against
#the alternative of spurious long memory.

Qutest=Qu.test(C02.month, 1+length(C02.month)^0.65, epsilon = 0.02)#0.65 value is between 0.5 and 0.8 and if length series >500 use epsilon =0.02
QUstatistical=Qutest[["W.stat"]]# value obtained by test
QUcritical=Qutest[["CriticalValues"]]# crtitical values
Qutestdecision=if(Qutest[["W.stat"]]<1.022){
  print("True long memory process")
} else {
  print("spurious long memory")
}
Qutestdecision

# Check this for information:https://www.jstor.org/stable/23243807?seq=1
##Nonparametric test for fractional cointegration (Nielsen (2010))
#stimation for fractional cointegration by Nielson (2010)
#FCI_N10(X, d1 = 0.1, m, mean_correct = c("mean", "init", "weighted",
#"none"), type = c("test", "rank"), alpha = 0.05)
#FCIN10test=FCI_N10(C02.month, d1 = 0.1, m=floor( 1+length(C02.month)^0.75), mean_correct = c("weighted"), type = c("test"), alpha = 0.05)# returns (null hypothesis: no fractional cointegration) or the estimated cointegrating rank.
#FCIN10statistic=FCIN10test$Ts;FCIN10statistic# test statistics
#FCIN10critical=FCIN10test$crit;FCIN10critical# critical value
#FCIN10testdecisionbyreject=FCIN10test$reject;FCIN10testdecisionbyreject# if false no fractional cointegration
#FCIN10testdecision=if(FCIN10test$Ts<FCIN10test$crit){
#  print("no fractional cointegration")
#} else {
#  print("there is fractional cointegration")
#}
#FCIN10testdecision

#Seasonal Autoegressive Fractional Integrated  moving average model

fit=auto.arima(C02.month,allowdrift = TRUE)
fit
acf(fit[["residuals"]],lag.max = 30)
fit[["fitted"]]
SARIMAmodel=print("ARIMA(5,0,0)(2,0,0)[12]") 
# The SARFIMA model aims to provide a general overview of the possible parameters to be considered within the model, which are considered as a starting point that are then indicated by trail and biases method to determine the parameters of the SARFIMA mdoel providing the best fit
fit1f=arfima(C02.month, order = c(2,0,0), numeach = c(1, 1), dmean = FALSE,
         whichopt = 0, itmean = FALSE, fixed = list(phi = NA, theta = NA,
                                                    frac = NA, seasonal = list(phi = NA, theta = NA, frac = NA), reg = NA),
         lmodel = c("d"), seasonal = list(order = c(1,0,2),
                                          period = 12, lmodel = c("d"), numeach = c(1, 1)),
         useC = 3, cpus = 1, rand = FALSE, numrand = NULL, seed = NA,
         eps3 = 0.01, xreg = NULL, reglist = list(regpar = NA, minn = -10,
                                                  maxx = 10, numeach = 1), check = F, autoweed = TRUE,
         weedeps = 0.05, adapt = TRUE, weedtype = c("B"),
         weedp = 2, quiet = FALSE, startfit = NULL, back = FALSE)

fit1f#coefficients
distance(fit1f)#compute distance between arfima modes
fit1f$z#observed
fit1f$n#. number of observed fitted database
fit1f$modes# contains alot of info,#muhat:The theoretical mean of the series before integration (if integer integration is done)

fit1f[["modes"]]

sarfimaar=as.data.frame(fit1f[["modes"]][[1]][["phi"]]);sarfimaar# ar parameter
sarfimama=as.data.frame(fit1f[["modes"]][[1]][["theta"]]);sarfimama#ma parameter
sarfimasar=as.data.frame(fit1f[["modes"]][[1]][["phiseas"]]);sarfimasar# seasonal ar parameter
sarfimasma=as.data.frame(fit1f[["modes"]][[1]][["thetaseas"]]);sarfimasma# seasonal ma parameter
#Check the following parameters they should be the same as mentioned before
fit1f[["modes"]][[1]][["phip"]]#same as ar parameter
fit1f[["modes"]][[1]][["thetap"]]# same as ma parameter
fit1f[["modes"]][[1]][["phiseasp"]]#same as seasonl ar parameter
fit1f[["modes"]][[1]][["thetaseasp"]]# same as seasonl ma parameter

sarfimad=as.data.frame(fit1f[["modes"]][[1]][["dfrac"]]);sarfimad# fractional integration
sarfimasd=as.data.frame(fit1f[["modes"]][[1]][["dfs"]]);sarfimasd#seasonal # fractional integration

sarfimaH=as.data.frame(fit1f[["modes"]][[1]][["H"]]);sarfimaH
sarfimaHs=as.data.frame(fit1f[["modes"]][[1]][["Hs"]]);sarfimaHs
sarfimaalpha=as.data.frame(fit1f[["modes"]][[1]][["alpha"]]);sarfimaalpha#FDWN noise
sarfimasalpha=as.data.frame(fit1f[["modes"]][[1]][["alphas"]]);sarfimasalpha#FDWN seasonal noise
sarfimadelta=as.data.frame(fit1f[["modes"]][[1]][["delta"]]);sarfimadelta#expected difference
sarfimaomega=as.data.frame(fit1f[["modes"]][[1]][["omega"]]);sarfimaomega#variacne intercept

sarfimafittedmean=as.data.frame(fit1f[["modes"]][[1]][["muHat"]]);sarfimafittedmean#fitted mean
sarfimaloglik=as.data.frame(fit1f[["modes"]][[1]][["loglik"]]);sarfimaloglik
sarfimapars=as.data.frame(fit1f[["modes"]][[1]][["pars"]]);sarfimapars#parameters used for optimization routine
sarfimaparsp=as.data.frame(fit1f[["modes"]][[1]][["parsp"]]);sarfimaparsp
sarfimase=as.data.frame(fit1f[["modes"]][[1]][["se"]]);sarfimase#standard errors
sarfimahess=as.data.frame(fit1f[["modes"]][[1]][["hess"]]);sarfimahess#passed to the hessian routine
sarfimafittedwithoutmean=as.data.frame(fit1f[["modes"]][[1]][["fitted"]]);sarfimafittedwithoutmean
sarfimaresiduals=as.data.frame(fit1f[["modes"]][[1]][["residuals"]]);sarfimaresiduals
sarfimasigma2=as.data.frame(fit1f[["modes"]][[1]][["sigma2"]]);sarfimasigma2


sarfimafitted=fit1f[["modes"]][[1]][["fitted"]]+fit1f[["modes"]][[1]][["muHat"]]
sarfimafitted
difforiginalwithfittedwithmean=C02.month-sarfimafitted
difforiginalwithfittedwithmean
plot(difforiginalwithfittedwithmean)
plot(sarfimafitted, type = "l", col = "red")
lines(C02.month, type = "l", col = "green")
sarfimafitted.df=as.data.frame(sarfimafitted);sarfimafitted.df
fit1fres=residuals(fit1f)
fit1fres#=fit1f[["modes"]][[1]][["residuals"]]
fit1f.fitted=fitted.values(fit1f);fit1f.fitted#=fit1f[["modes"]][[1]][["fitted"]]

#The approximate or (almost) exact Fisher information matrix of an ARFIMA process
fit1f # provide the coefficient that will be used in the following command
iARFIMA(phi=0.512,theta=0,phiseas = 0.45,thetaseas = 0.97,period = 12,dfrac = TRUE,dfs=TRUE,exact=TRUE)

# Rsquare value

df.fit1fobserved = as.data.frame(fit1f$z)
df.fit1fobserved
fit1fres.df=as.data.frame(fit1fres)
fit1fres.df


df.fit1ffitted.R2 = df.fit1fobserved+fit1fres.df
df.fit1ffitted.R2 

df.fit1ffitted.R2ts=ts(df.fit1ffitted.R2)
df.fit1fobservedts=ts(df.fit1fobserved)


df.fit1ffittedmean.R2ts=ts(sarfimafitted)
df.fit1fobservedts=ts(df.fit1fobserved)
R2SARFIMAfittedwithmean=rsq(glm(df.fit1fobservedts~df.fit1ffittedmean.R2ts),adj = FALSE)
R2SARFIMAfittedwithmean
#######################################

#CHECK:
#check the residuals of the fitted model:

#Pvalue< 0.05 and Acf should be embedded in the boundary  
# it is common to find that the LB test fails on residuals even though the forecasts might be good using Pvalue<0.05 as a threshold
#if it is > 0.05 then it suggests that there is a little more information in the data than is captured in the model.
#lag=For seasonal time series, use h=  min(  2 m ,  n/5).;m= is the period of seasonality & where n is the length of the series.

acf(fit1fres.df, lag.max = 50)  #check ACF
sarfimaacfresid=acf(fit1f[["modes"]][[1]][["residuals"]],lag.max = 50)
sarfimaacfresiduals=as.data.frame(sarfimaacfresid$acf,sarfimaacfresid$lag);sarfimaacfresiduals

sarfimaacffittedmodel=acf(fit1f[["modes"]][[1]][["fitted"]],lag.max = 30)
sarfimaacfrealfitted=as.data.frame(sarfimaacffittedmodel$acf,sarfimaacffittedmodel$lag);sarfimaacfrealfitted

#Ljung Box test

#is a statistical test that checks if autocorrelation exists in a time series and check the prediction models
#Ideally, we would like to fail to reject the null hypothesis. That is, we would like to see the p-value of the test be greater than 0.05 because 
#this means the residuals for our time series model are independent, which is often an assumption we make when creating a model.
#if Pvalue is much larger than 0.05. Thus, we fail to reject the null hypothesis of the test and conclude that the data values are independent.

LJB=Box.test.lagvalue=min(24,length(C02.month)/5)# the value of the lag component used in the BOX.TEST
LJB

fit1f[["modes"]][[1]][["phi"]]# ar parameter
fit1f[["modes"]][[1]][["theta"]]#ma parameter
#fitdf= p+q
BT.C=Box.test(sarfimaresiduals,lag=24,type ="Ljung-Box",fitdf = 1) # check Pvalue 
BT.C
BT.C$p.value# if it is > 0.05 then it is accepted check script 401
BT.C$statistic

BT.Cpvalue=c(BT.C$p.value)
BT.Cpvalue
BT.Cstatistic=c(BT.C$statistic)
BT.Cstatistic
Criteria=c(0.05)
Criteria
BT.Cresult=ifelse(BT.Cpvalue>Criteria ,"ljung Box test result failed to reject the null hypothesis of the test but it is not confirmed and conclude that the data values are independent because they are not dependant",
                   "ljung Box test rejected null hypothesis.Thus the time series is showing dependence ")
BT.Cresult

###Froeacsting
##Prediction: This is performed to visualize the confidence 95% interval "Limits"


predict=predict(fit1f, n.ahead = 48, prop.use = "default",
        newxreg = NULL, predint = 0.95, exact = c("default", T, F),
        setmuhat0 = FALSE, cpus = 1, trend = NULL, n.use = NULL,
        xreg = NULL)
predict
sarfimapredict=as.data.frame(predict[[1]][["Forecast"]]);sarfimapredict
sarfimaexactsd=as.data.frame(predict[[1]][["exactSD"]]);sarfimaexactsd
sarfimalimitsd=as.data.frame(predict[[1]][["limitSD"]]);sarfimalimitsd

predict[[1]][["SDForecasts"]]# je ne trouve pas un neccessarity del'utilize
sarfimasigma2=as.data.frame(predict[[1]][["sigma2"]]);sarfimasigma2

sarfimaexactvar=as.data.frame(predict[[1]][["exactVar"]]);sarfimaexactvar
sarfimalimitvar=as.data.frame(predict[[1]][["limitVar"]]);sarfimalimitvar


plot(predict[[1]][["exactVar"]])
plot(predict[[1]][["exactSD"]])
plot(predict[[1]][["Forecast"]])

predictforecast=as.data.frame(predict[[1]][["Forecast"]])
predictSD=as.data.frame(predict[[1]][["SDForecasts"]])
predictsigma2=as.data.frame(predict[[1]][["sigma2"]])
predictexactvar=as.data.frame(predict[[1]][["exactVar"]])
predictlimitvar=as.data.frame(predict[[1]][["limitVar"]])
predictexactSD=as.data.frame(predict[[1]][["exactSD"]])
predictlimitSD=as.data.frame(predict[[1]][["limitSD"]])

plot(predict)
show(predict)
predict$z#observed data by the arfima model
predict$name#model fitted to this prediction
sarfimapredictconfidence=as.data.frame(predict$predint);sarfimapredictconfidence#limit of the prediction condifence

###Forecasting
#forecast=forecast(fit1f$z,h=48)
#forecast
#forecastresiduals=as.data.frame(forecast[["residuals"]]);forecastresiduals

#forecastacfresiduals=acf(forecastresiduals,lag.max = 30)
#forecastacfresid=as.data.frame(forecastacfresiduals$acf,forecastacfresiduals$lag);forecastacfresid

#forecast[["model"]][["fitted"]]
#forecast$x#=original database
#forecastmodelerror=C02.month-forecast[["model"]][["fitted"]]
#forecastmodelerror
#forecasterrorinfittedmodel=as.data.frame(forecastmodelerror);forecasterrorinfittedmodel
#plot(forecastmodelerror)
#forecast$model#print input in a line script below
#forecastmodel=print(" Smoothing parameters:
#    alpha = 0.5789 
#    gamma = 1e-04 

#  Initial states:
#   l = 70.3898 
#   s = 1.775 0.959 0.4496 0.2497 0.2148 0.2749
#          0.4512 0.7596 1.1605 1.6705 1.8675 2.1678

# sigma:  0.6573

#    AIC     AICc      BIC 
#23023.07 23023.37 23103.93 ")

#plot(forecast)
#df.forecast = as.data.frame(forecast)
#df.forecast
#sarfima.forecast = df.forecast$`Point Forecast`#forecasted results
#sarfima.forecast
#sarfima.Low.80 = df.forecast$`Lo 80`
#sarfima.High.80= df.forecast$`Hi 80`
#sarfima.Low.95 = df.forecast$`Lo 95`
#sarfima.High.95= df.forecast$`Hi 95`
#plot(sarfima.forecast)

#diffinforecast=as.data.frame(sarfima.forecast)-sarfimapredict#difference between using ARFIMA prediction and forecasting
#diffinforecast
#ts.plot(diffinforecast)


###Extracts the tacvfs of a fitted object
#Extracts the theoretical autocovariance functions (tacvfs) from
#a fitted arfima or one of its modes (an ARFIMA) object.
arfima.tacvf=tacvf(fit1f, xmaxlag = 30, forPred = FALSE, tacf = TRUE, n.ahead=0)
plot(arfima.tacvf)


##Plots the theoretical autocorralation functions (tacfs) of one or more fits.
#Plots the theoretical autocorralation functions (tacfs) of one or more fits.
tacfresult=tacfplot(fit1f)
plot(arfima.tacvf,tacf=TRUE)
vcov(fit1f)#returns a covariance variance matrix of the arfima model

#Dirctional accuracy test

#df.existedDAT = as.data.frame(fit1f$z[437:466])
#df.existedDAT
#df.forecastDAT = as.data.frame(predict[[1]][["Forecast"]])
#df.forecastDAT#=predictforecast function
#DACTest(,,test = c("AG"), conf.level = 0.95)
#DACTest(,,test = c("PT"), conf.level = 0.95)
#sarfimapredict
#sarfima.forecast

## Autocorrelation of the residuals forecasted results
acf(forecast[["residuals"]], lag.max = 30)  #check ACF
#Ljung Box test

#is a statistical test that checks if autocorrelation exists in a time series and check the prediction models
#Ideally, we would like to fail to reject the null hypothesis. That is, we would like to see the p-value of the test be greater than 0.05 because 
#this means the residuals for our time series model are independent, which is often an assumption we make when creating a model.
#if Pvalue is much larger than 0.05. Thus, we fail to reject the null hypothesis of the test and conclude that the data values are independent.

LJB=Box.test.lagvalue=min(24,length(C02.month)/5)# the value of the lag component used in the BOX.TEST
LJB

fit1f[["modes"]][[1]][["phi"]]# ar parameter
fit1f[["modes"]][[1]][["theta"]]#ma parameter
#fitdf= p+q
#BT.Cf=Box.test(forecast[["residuals"]],lag=24,type ="Ljung-Box",fitdf = 0) # check Pvalue 
#BT.Cf
#BT.Cf$p.value# if it is > 0.05 then it is accepted check script 401
#BT.Cf$statistic

#BT.Cfpvalue=c(BT.Cf$p.value)
#BT.Cfpvalue
#BT.Cfstatistic=c(BT.Cf$statistic)
#BT.Cfstatistic
#Criteria=c(0.05)
#Criteria
#BT.Cfresult=ifelse(BT.Cfpvalue>Criteria ,"ljung Box test result failed to reject the null hypothesis of the test but it is not confirmed and conclude that the data values are independent because they are not dependant",
#                  "ljung Box test rejected null hypothesis.Thus the time series is showing dependence ")
#BT.Cfresult

#######################################

###Accuracy parameters

df.fit1ffitted.R2 #observed + residuals
as.data.frame(sarfimafitted)#mean + fitted of the model
diffinfitting=df.fit1ffitted.R2-as.data.frame(sarfimafitted)
diffinfitting
ts.plot(diffinfitting)

#Accuracy of the fitted model
accuracysarfima=accuracy(ts(df.fit1ffitted.R2),ts(fit1f$z),test= NULL, d=0,D=1)
accuracysarfima#fitted by mean
accuracysarfima.df=as.data.frame(accuracysarfima);accuracysarfima.df

accuracysarfimaresidfitted=accuracy(ts(df.fit1ffitted.R2),ts(fit1f$z),test= NULL, d=0,D=1)
accuracysarfimaresidfitted#fitted by residuals+observed data
accuracysarfimaresidfitted.df=as.data.frame(accuracysarfimaresidfitted);accuracysarfimaresidfitted.df
#Accuracy of the forecast
#accuracyforecast=accuracy(forecast)
#accuracyforecast
#accuracyforecast.df=as.data.frame(accuracyforecast);accuracyforecast.df


library(Metrics)

# df.fit1ffitted.R2
# fit1f$z
fit1fz.df=as.data.frame(fit1f$z)
fit1fz.df

absoluteerror=ae(df.fit1ffitted.R2,fit1fz.df);absoluteerror
absolutepercenterror=ape(df.fit1ffitted.R2,fit1fz.df);absolutepercenterror
auc(df.fit1ffitted.R2,fit1fz.df)#e area under the ROC curve is equal to the probability that a randomly chosen positive observation has a higher predicted value than a randomly chosen negative value. In order to compute this probability, we can calculate the Mann-Whitney U statistic. This method is very fast
averageactualbiggerthanpredicted=bias(ts(df.fit1ffitted.R2),ts(fit1fz.df))#bias computes the average amount by which actual is greater than predicted
averageactualpercentbiggerthanpredicted=percent_bias(ts(df.fit1ffitted.R2),ts(fit1fz.df));averageactualpercentbiggerthanpredicted#bias computes the average amount by which actual is greater than predicted

meanabsoluteerror=mae(ts(df.fit1ffitted.R2),ts(fit1fz.df));meanabsoluteerror
meanabsoluteprecenterror=mape(ts(df.fit1ffitted.R2),ts(fit1fz.df));meanabsoluteprecenterror
meanabsolutescaleserror=mase(ts(df.fit1ffitted.R2),ts(fit1fz.df));meanabsolutescaleserror
medianabsoluteerror=mdae(ts(df.fit1ffitted.R2),ts(fit1fz.df));medianabsoluteerror
meanquadraticweightedkappa=MeanQuadraticWeightedKappa(ts(df.fit1ffitted.R2),ts(fit1fz.df));meanquadraticweightedkappa
meansquarederror=mse(ts(df.fit1ffitted.R2),ts(fit1fz.df));meansquarederror
meansquaredlogerror=msle(ts(df.fit1ffitted.R2),ts(fit1fz.df));meansquaredlogerror
relativeabsoluteerror=rae(ts(df.fit1ffitted.R2),ts(fit1fz.df));relativeabsoluteerror
rootmeansquareerror=rmse(ts(df.fit1ffitted.R2),ts(fit1fz.df));rootmeansquareerror
rootmeansquarelogerror=rmsle(ts(df.fit1ffitted.R2),ts(fit1fz.df));rootmeansquarelogerror
rootrelativesquarederror=rrse(ts(df.fit1ffitted.R2),ts(fit1fz.df));rootrelativesquarederror
relativesquarederror=rse(ts(df.fit1ffitted.R2),ts(fit1fz.df));relativesquarederror
squarederror=se(ts(df.fit1ffitted.R2),ts(fit1fz.df));squarederror
squaredlogerror=sle(ts(df.fit1ffitted.R2),ts(fit1fz.df));squaredlogerror
SymmetricMeanAbsolutePercentageError=smape(ts(df.fit1ffitted.R2),ts(fit1fz.df));SymmetricMeanAbsolutePercentageError
sumofsquarederrors=sse(ts(df.fit1ffitted.R2),ts(fit1fz.df));sumofsquarederrors

plot(months.ratio)

#####################################################


#Shifting technique
library(dplyr)

C02.df <- data.frame(
  Year = floor(time(C02.month)),
  Month = cycle(C02.month),
  Value = as.numeric(C02.month)
)

# Aggregate data by year and calculate variance
yearly_variance <- 11/12*aggregate(Value ~ Year, data = C02.df, FUN= var)# 11/12: (n-1)/n to provide the monthly population variance
yearly_mean <- aggregate(Value ~ Year, data = C02.df, FUN = mean)

yearly_variance=data.frame(Year= seq(1883, 2018), Value=yearly_variance$Value)

head(yearly_variance)


# Calculate the overall variance

overall_variance <- mean(yearly_variance$Value[1:135])

variance_threshold <- 2 * overall_variance
shifting_years <- yearly_variance$Year[yearly_variance$Value > variance_threshold]

sarfimafittedwithoutmean
yearly_mean <- aggregate(Value ~ Year, data = C02.df, FUN = mean)

# Create a logical vector indicating shifting years

shifting_years_logical <- C02.df$Year %in% shifting_years


# Calculate the constant mean for non-shifting years

mean_non_shifting <- mean(C02.df$Value[!shifting_years_logical])

# Calculate the varying mean for shifting years

mean_shifting <- aggregate(Value ~ Year, data = C02.df[shifting_years_logical, ], FUN = mean)

MS_SARFIMA_df <- data.frame(
  Year = C02.df$Year,
  MS_SARFIMA = sarfimafittedwithoutmean + ifelse(shifting_years_logical, mean_shifting$Value[match(C02.df$Year, mean_shifting$Year)], mean_non_shifting)
)
MS_SARFIMA_df

##########################

# Create a subset of MS_SARFIMA_df for the shifting years

MS_SARFIMA_shifting_df<- MS_SARFIMA_df[MS_SARFIMA_df$Year %in% shifting_years, ]


# Calculate the variance of MS-SARFIMA for each shifting year

yearly_variance_MS_SARFIMA_shifting <- 11/12*aggregate(x ~ Year, data = MS_SARFIMA_shifting_df, FUN = var)
yearly_variance_MS_SARFIMA_shifting=data.frame(shifting_years[1:15],yearly_variance_MS_SARFIMA_shifting$x)
colnames(yearly_variance_MS_SARFIMA_shifting) <- c("Year", "YearlyVariance_MS_SARFIMA_Shifting")
yearly_variance_MS_SARFIMA_shifting

# Calculate the mean of MS-SARFIMA for each shifting year

yearly_mean_shifting <- aggregate(Value ~ Year, data = C02.df[C02.df$Year %in% shifting_years, ], FUN = mean)

########### Calculate the shifting values for each shifting year

shifting_values <- numeric(length(yearly_variance_MS_SARFIMA_shifting$YearlyVariance_MS_SARFIMA_Shifting))


for (i in seq_along(shifting_values)) {
  if (overall_variance > yearly_variance_MS_SARFIMA_shifting$YearlyVariance_MS_SARFIMA_Shifting[i]) {
    shifting_values[i] <- (overall_variance / yearly_variance_MS_SARFIMA_shifting$YearlyVariance_MS_SARFIMA_Shifting[i]) * yearly_mean_shifting$Value[i]
  } else {
    shifting_values[i] <- (yearly_variance_MS_SARFIMA_shifting$YearlyVariance_MS_SARFIMA_Shifting[i] / overall_variance) * yearly_mean_shifting$Value[i]
  }
}

shifting_values_df <- data.frame(
  Year = yearly_variance_MS_SARFIMA_shifting$Year,
  ShiftingValues = shifting_values
)

shifting_values_df

################################################
# Visualize the seasonal effect detected in specefic months starting from 1883 

data.janv=window(C02.month, start=c(1883,1),frequency=TRUE)
data.janv
var(data.janv)
data.fev=window(C02.month, start=c(1883,2),frequency=TRUE)
data.fev
data.mars=window(C02.month, start=c(1883,3),frequency=TRUE)
data.mars
data.avril=window(C02.month, start=c(1883,4),frequency=TRUE)
data.avril
data.mai=window(C02.month, start=c(1883,5),frequency=TRUE)
data.mai
data.june=window(C02.month, start=c(1883,6),frequency=TRUE)
data.june
data.juillet=window(C02.month, start=c(1883,7),frequency=TRUE)
data.juillet
data.aout=window(C02.month, start=c(1883,8),frequency=TRUE)
data.aout
data.septembre=window(C02.month, start=c(1883,9),frequency=TRUE)
data.septembre
data.octobre=window(C02.month, start=c(1883,10),frequency=TRUE)
data.octobre
data.novembre=window(C02.month, start=c(1883,11),frequency=TRUE)
data.novembre
data.decembre=window(C02.month, start=c(1883,12),frequency=TRUE)
data.decembre

plot(data.janv)
m1=mean(data.janv)
v1=var(data.janv)

#high mean and variance
plot(data.fev)
m2=mean(data.fev)
v2=var(data.fev)
#high mean and variance
plot(data.mars)#
m3=mean(data.mars)
v3=var(data.mars)
# low mean and variance
plot(data.avril)#
m4=mean(data.avril)
v4=var(data.avril)
# low mean and variance
plot(data.mai)#
m5=mean(data.mai)
v5=var(data.mai)
# low mean and variance
plot(data.june)# 
m6=mean(data.june)
v6=var(data.june)
# low mean and variance
plot(data.juillet)#
m7=mean(data.juillet)
v7=var(data.juillet)
# low mean and variance
plot(data.aout)# 
m8=mean(data.aout)
v8=var(data.aout)
# low mean and variance
plot(data.septembre)#
m9=mean(data.septembre)
v9=var(data.septembre)
# low mean and variance
plot(data.octobre)#
m10=mean(data.octobre)
v10=var(data.octobre)
# low mean and variance
plot(data.novembre)# 
m11=mean(data.novembre)
v11=var(data.novembre)
# low mean and variance
plot(data.decembre)#
m12=mean(data.decembre)
v12=var(data.decembre)


monthsmeanvalues=data.frame(m1,m2,m3,m4,m5,m6,m7,m8,m9,m10,m11,m12)
monthsvariancevalues=data.frame(v1,v2,v3,v4,v5,v6,v7,v8,v9,v10,v11,v12)


plot(months.ratio)
plot(months.ratiovar)

################################################
# Seasonal Markov switching component

##Monthly seasonal effect

# Influenced-shifting states

# Viewing the variance 
vr1=((v1/overall_variance)*100)
vr2=((v2/overall_variance)*100)
vr3=((v3/overall_variance)*100)
vr4=((v4/overall_variance)*100)
vr5=((v5/overall_variance)*100)
vr6=((v6/overall_variance)*100)
vr7=((v7/overall_variance)*100)
vr8=((v8/overall_variance)*100)
vr9=((v9/overall_variance)*100)
vr10=((v10/overall_variance)*100)
vr11=((v11/overall_variance)*100)
vr12=((v12/overall_variance)*100)



# the postive factor within the months.ratiovar conclude that jan, fev et dec are months that must be taken into account 
months.ratio

vm=array(data = c(months.ratio["janv.ratio",1]
,months.ratio["fev.ratio",1]
,months.ratio["mars.ratio",1]
,months.ratio["avril.ratio",1]
,months.ratio["mai.ratio",1]
,months.ratio["june.ratio",1]
,months.ratio["juillet.ratio",1]
,months.ratio["aout.ratio",1]
,months.ratio["septembre.ratio",1]
,months.ratio["novembre.ratio",1]
,months.ratio["decembre.ratio",1]
))


moi=as.numeric(0)
for (i in 1:12) {
  if (vm[i]>=1.2) {
    moi[i]=vm[i]
  }
}
moi# influenced*-shifting states: months that the mean value considerding the overall mean value: by this mars is also considered as an impacting factor 
#The results indicate that Janvier, fev, mars, decembre are considerd as influenced-shifitng months in this study case following a 80% C.I

shifting_years=shifting_years[1:15]

# Filter the dataset for the shifting years and the specified months (Jan, Feb, Mar, Dec)
subset_df <- C02.df[C02.df$Year %in% shifting_years & C02.df$Month %in% c(1, 2, 3, 12), ]

sarfima.df <- data.frame(
  Year = floor(time(sarfimafitted)),
  Month = cycle(sarfimafitted),
  Value = as.numeric(sarfimafitted)
)

subset_sarfima <- sarfima.df[sarfima.df$Year %in% shifting_years & sarfima.df$Month %in% c(1,2,3,12),]
Diff.df=subset_df$Value-subset_sarfima$Value
mv=sarfimafittedmean$`fit1f[["modes"]][[1]][["muHat"]]`


subset_df$Transitionstaterequirement <- ifelse(abs(subset_df$Value - subset_sarfima$Value) > mv, 1, 0)
sum_states <- sum(subset_df$Transitionstaterequirement);sum_states

#############################################

###################Limite of the severe values following the shifting-influenced months
moi# shifting-influenced months: janvier,fevrier,mars, decembre

C02.janv
C02.fev
C02.mars
C02.decembre
#janv
janvquartile3 <- quantile(C02.janv, 0.75)
janvquartile1 <- quantile(C02.janv, 0.25)
janvq=janvquartile3+1.5*(janvquartile3-janvquartile1)
#fev
fevquartile3 <- quantile(C02.fev, 0.75)
fevquartile1 <- quantile(C02.fev, 0.25)
fevq=fevquartile3+1.5*(fevquartile3-fevquartile1)
#Mars
Marsquartile3 <- quantile(C02.mars, 0.75)
Marsquartile1 <- quantile(C02.mars, 0.25)
Marsq=Marsquartile3+1.5*(Marsquartile3-Marsquartile1)
#Decembre
Decquartile3 <- quantile(C02.decembre, 0.75)
Decquartile1 <- quantile(C02.decembre, 0.25)
Decq=Decquartile3+1.5*(Decquartile3-Decquartile1)
extremelimit=(janvq+fevq+Marsq+Decq)/4

####################################

#####Create a hidden Markov model for seasonal probabilistic 
shifting_years

library(HMM)
# Important note
# The results of the subset_df$Transitionstaterequirement are exported to view the transition from the states to provide the transition probability since it is easier to view
# Transition probability matrix is calibrated based on the Months distribution effect
# The values probided are the output of a counting proces of the states within the shifting_years 

HMMS=initHMM(c("January","February","March","December"),c("January","February","March","December"),startProbs = matrix(c(1,0,0,0)),transProbs=matrix(c(0,9.52,9.52,9.52,
                                                                                                                                                       23.8,4.76,0,4.76,
                                                                                                                                                       0,9.52,4.76,4.76,
                                                                                                                                                       9.52,4.76,4.76,0),4,4))
HMMS.df <- as.data.frame(HMMS)
HMMS.df
HMMS$States#Vector with the names of the states.
HMMS$Symbols#Vector with the names of the symbols.
HMMS$startProbs#Annotated vector with the starting probabilities of the states.
#States transition probabilistic matrix based on the occurance of the Seasonal probabilistic transition matrix
HMMS$transProbs#Annotated matrix containing the transition probabilities between the states.
#Seasonality probabilistic transition matrix
HMMS$emissionProbs#Annotated matrix containing the emission probabilities of the states.
HMMSstates.df=as.data.frame(HMMS$States)
write.xlsx(
  HMMSstates.df,
  file="Seasonal Markov chain.xlsx",
  sheetName = "Months",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
HMMSsymbol.df=as.data.frame(HMMS$Symbols)
write.xlsx(
  HMMSsymbol.df,
  file="Seasonal Markov chain.xlsx",
  sheetName = "Conditions",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
HMMSStartprobs.df=as.data.frame(HMMS$startProbs)
write.xlsx(
  HMMSStartprobs.df,
  file="Seasonal Markov chain.xlsx",
  sheetName = "Start probabilistic",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
HMMStransprobs.df=as.data.frame(HMMS$transProbs)
write.xlsx(
  HMMStransprobs.df,
  file="Seasonal Markov chain.xlsx",
  sheetName = "Transition probabilistic Matrix between months",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
HMMSemissionprobs.df=as.data.frame(HMMS$emissionProbs)
write.xlsx(
  HMMSemissionprobs.df,
  file="Seasonal Markov chain.xlsx",
  sheetName = "Seasonality occurance probability matrix",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

#####################

##Simulates a path of states and observations 
HMMSpath=simHMM(HMMS, sum_states)
HMMSpath$states# States within the year of shifting 
HMMSpath$observation
####################
# Method 1:
##Simulates a path of states and observations 
HMMSpaths <- matrix(NA, nrow = sum_states, ncol = 100000)
for (i in 1:100000) {
  HMMSpaths[, i] <- simHMM(HMMS, sum_states)$observation  
}

most_frequent_months <- character(sum_states)

# Loop through each row (position in the sequence) to find the most frequent month
for (i in 1:sum_states) {
  freq_table <- table(HMMSpaths[i, ])
    if (length(freq_table) > 0) {
    most_frequent_months[i] <- names(which.max(freq_table))
  } else {
    most_frequent_months[i] <- NA
  }
}

# Display the most frequent month for each position
most_frequent_months

#Method 2: Performed in python
# IF condition with loop ( Choose the simulation (most_frequent_months) with coefficient of variation >=80 ):
subset_df$Transitionstaterequirement=c(1,1,0,0,0,0,0,0,0,1,0,0,0,1,0,1,0,0,1,0,0,1,0,0,0,0,0,1,1,1,1,0,0,0,1,0,1,1,1,0,0,0,0,1,1,1,0,0,0,0,0,1,1,0,1,0,1,1,0,0)
############################

shifting_values
replicatedshifting_values <- rep(shifting_values, each = 4)
shifting_years
subset_df$Transitionstaterequirement
dataframecms=data.frame(subset_df,subset_sarfima$Value,replicatedshifting_values)
subset_sarfima$Value

dataframecms
dataframecms$cmssarfima <- ifelse(
  dataframecms$Transitionstaterequirement == 1, 
  dataframecms$subset_sarfima.Value + dataframecms$replicatedshifting_values, 
  dataframecms$subset_sarfima.Value
)
# CmsSARFIMA model as computed is for the shifting_years. Manually they will be updated with all the other timesteps (Years of not shifting) which will have the same values as the sarfima model.  
###################################################
#######################################################################
###Plots empirical expected number of upcrossings level u with P(Y<u)

Time <- read_excel("C:/Users/,,,,,,,,,,,,,,,,,,,,,,,,,,,,,.xlsx",range = cell_cols("A") )
dataset <- read_excel("C:/Users/,,,,,,,,,,,,,,,,,,,,,,,,.xlsx",range = cell_cols("B") )
sarfima <- read_excel("C:/Users/,,,,,,,,,,,,,,,,,,,,,,,.xlsx",range = cell_cols("C") )
hcmssarfima <- read_excel("C:/Users/,,,,,,,,,,,,,,,,,,,,,,,.xlsx",range = cell_cols("E") )

#ENu_graph function Plots empirical expected number of upcrossings of 
#level u with respect to P(Y<u)

library(NHMSAR)
uo= seq(min(dataset$Dataset),max(dataset$Dataset),by=10)
um= seq(min(sarfima$SARFIMA),max(sarfima$SARFIMA),by=10)
uhcms= seq(min(hcmssarfima$`HC-MS-SARFIMA`),max(hcmssarfima$`HC-MS-SARFIMA`),by=10)

#Choose bounds of the input and write the u function as a sequence based 
#on the sequential distribution of the dataset, and the simulations 
u=seq(0,490,by=10)

# Dataset empirical expected number of upcrossings level
EENUO = ENu_graph(ts(dataset),u)#probability of data<ulog
EENUO$u#sequence of levels
EENUO$F#empirical cdf: P(data<u)
EENUO$Nu#number of upcrossings
EENUO$CI#fluctuation interval
EENUOdf=as.data.frame(EENUO)
EENUOdf
##################################

# MSAR empirical expected number of upcrossings level
EENUM = ENu_graph(ts(sarfima),u)#probability of data<ulog
EENUM$u#sequence of levels
EENUM$F#empirical cdf: P(data<u)
EENUM$Nu#number of upcrossings
EENUM$CI#fluctuation interval
EENUMdf=as.data.frame(EENUM)
EENUMdf

##################################

# SMSAR empirical expected number of upcrossings level
EENUS = ENu_graph(ts(hcmssarfima),u)#probability of data<ulog
EENUS$u#sequence of levels
EENUS$F#empirical cdf: P(data<u)
EENUS$Nu#number of upcrossings
EENUS$CI#fluctuation interval
EENUSdf=as.data.frame(EENUS)
EENUSdf

write.xlsx(
  EENUOdf,
  file="numberofupcrossing.xlsx",
  sheetName = "Dataset",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE
)
write.xlsx(
  EENUMdf,
  file="numberofupcrossing.xlsx",
  sheetName = "SARFIMA",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE
)
write.xlsx(
  EENUSdf,
  file="numberofupcrossing.xlsx",
  sheetName = "hcmssarfima",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE
)


################################################################################
##Extracrtion
#original.dataframe[original.dataframe>=extremelimit]
#filter(original.dataframe,original>=extremelimit)
#original.dataframe %>% filter(original >=extremelimit)
Time.dataframe=data.frame(Time)
original.dataframe=data.frame(dataset)
sarfima.dataframe=data.frame(sarfima)
hcmssarfima.dataframe=data.frame(hcmssarfima)


########### Severevalues
#???The naming 250 is not the extreme limit 
original250=filter(original.dataframe,dataset>=extremelimit);original250
#Index
original.dataframeindex250 <- which(original.dataframe$Dataset >= extremelimit);original.dataframeindex250
original.dataframeindex250.df=data.frame(original.dataframeindex250);original.dataframeindex250.df
#Date based on the Index
Dataset.dataframe250=original.dataframe$Dataset[original.dataframeindex250];Dataset.dataframe250
Time.dataframe250=Time.dataframe$Time[original.dataframeindex250];Time.dataframe250
#Creating the severe File
Originalseverevalue250=data.frame(Time.dataframe250,Dataset.dataframe250)
colnames(Originalseverevalue250) <- c("Date", "Dataset269");Originalseverevalue250
#Saving
write.xlsx(
  Originalseverevalue250,
  file="SevereValues269.xlsx",
  sheetName = "Dataset",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE
)

#severevalues
sarfima250=filter(sarfima.dataframe,sarfima>=extremelimit);sarfima250
#Index
sarfima.dataframeindex250 <- which(sarfima.dataframe$SARFIMA >= extremelimit);sarfima.dataframeindex250
sarfima.dataframeindex250.df=data.frame(sarfima.dataframeindex250);sarfima.dataframeindex250.df
#Date based on the Index
sarfima.dataframe250=sarfima.dataframe$SARFIMA[sarfima.dataframeindex250];sarfima.dataframe250
Time.dataframe250=Time.dataframe$Time[sarfima.dataframeindex250];Time.dataframe250
#Creating the severe File
sarfimaseverevalue250=data.frame(Time.dataframe250,sarfima.dataframe250)
colnames(sarfimaseverevalue250) <- c("Date", "sarfima269");sarfimaseverevalue250
#Saving
write.xlsx(
  sarfimaseverevalue250,
  file="SevereValues269.xlsx",
  sheetName = "sarfima",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE
)



#severevalues
hcmssarfima250=filter(hcmssarfima.dataframe,hcmssarfima>=extremelimit);hcmssarfima250
#Index
hcmssarfima.dataframeindex250 <- which(hcmssarfima.dataframe$HC.MS.SARFIMA >= extremelimit);hcmssarfima.dataframeindex250
hcmssarfima.dataframeindex250.df=data.frame(hcmssarfima.dataframeindex250);hcmssarfima.dataframeindex250.df
#Date based on the Index
hcmssarfima.dataframe250=hcmssarfima.dataframe$HC.MS.SARFIMA[hcmssarfima.dataframeindex250];hcmssarfima.dataframe250
Time.dataframe250=Time.dataframe$Time[hcmssarfima.dataframeindex250];Time.dataframe250
#Creating the severe File
hcmssarfimaseverevalue250=data.frame(Time.dataframe250,hcmssarfima.dataframe250)
colnames(hcmssarfimaseverevalue250) <- c("Date", "hcmssarfima269");hcmssarfimaseverevalue250
#Saving
write.xlsx(
  hcmssarfimaseverevalue250,
  file="SevereValues269.xlsx",
  sheetName = "hcmssarfima",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE
)





######################################

do=density(dataset$Dataset,n=12,from = 0,to = 480)
ds=density(sarfima$SARFIMA,n=12,from = 0,to = 480)
dms=density(hcmssarfima$`HC-MS-SARFIMA`,n=12,from = 0,to = 480)
plot(do)
lines(ds)
lines(dms)
do$x
do$y# density
denistyframeoriginal=data.frame(do$x,do$y)
denistyframesarfima=data.frame(ds$x,ds$y)
denistyframehcmssarfima=data.frame(dms$x,dms$y)


write.xlsx(
  denistyframeoriginal,
  file="Kerneldistribution.xlsx",
  sheetName = "Dataset",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE
)
write.xlsx(
  denistyframesarfima,
  file="Kerneldistribution.xlsx",
  sheetName = "SARFIMA",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE
)
write.xlsx(
  denistyframehcmssarfima,
  file="Kerneldistribution.xlsx",
  sheetName = "HCMSSARFIMA",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE
)

################################
#Prediction

#open the file with the dataset and simualtions that are determined for the period in the article 
datasetp <- read_excel("C:/Users/,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,.xlsx",range = cell_cols("B") )
sarfimap <- read_excel("C:/Users/,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,.xlsx",range = cell_cols("C") )
hcmssarfimap <- read_excel("C:/Users/B,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,.xlsx",range = cell_cols("D") )
forecastfit <- read_excel("C:/Users/,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,.xlsx",range = cell_cols("I") )


#Determine the order of the model
acf(hcmssarfimap)
pacf(hcmssarfimap)
length(hcmssarfimap$HCMSSARFIMA)
x=hcmssarfimap


acf(x)
pacf(x)

#Fitting the model
NN=nnetar(ts(x),p=14,P=2,repeats = 20)
NN$fitted
NN$residuals
NN$method#(x;y) 2 hidden layers each one has the value of neurons wrtten
NN$model # model information
summary(NN)

# Forecasting the NN
NNforecast=forecast(NN, h=25)
NNforecastdf=data.frame(NNforecast)
plot(forecastfit$forecastfit)
lines(NNforecastdf$Point.Forecast)

plot(NNforecastdf$Point.Forecast)
lines(forecastfit$forecastfit)
# Accuracy
NNaccuracy=accuracy(NNforecast)
NNaccuracydf=data.frame(NNaccuracy)
# Dataframe of the results
NNfitteddf=data.frame(NN$fitted)
NNresidualsdf=data.frame(NN$residuals)
NNmethoddf=data.frame(NN$method)
NN$model

NNmodeldf=data.frame("Average of 20 networks, each of which is
a 14-8-1 network with 129 weights options were - linear output units")
# Saving the results
write.xlsx(
  NNfitteddf,
  file="NN-MSSARFIMA-.xlsx",
  sheetName = "Fitted Model",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE
)
write.xlsx(
  NNresidualsdf,
  file="NN-MSSARFIMA-.xlsx",
  sheetName = "Residuals",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE
)
write.xlsx(
  NNmethoddf,
  file="NN-MSSARFIMA-.xlsx",
  sheetName = "Method",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE
)
write.xlsx(
  NNmodeldf,
  file="NN-MSSARFIMA-.xlsx",
  sheetName = "Model",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE
)
write.xlsx(
  NNforecastdf,
  file="NN-MSSARFIMA-.xlsx",
  sheetName = "Forecast",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE
)
write.xlsx(
  NNaccuracydf,
  file="NN-MSSARFIMA-.xlsx",
  sheetName = "Accuracy",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE
)

###################################


#Determine the order
acf(datasetp)
pacf(datasetp)
x=datasetp$Dataset


acf(x)
pacf(x)

#Fitting the model
NN=nnetar(ts(x),p=11,P=1,repeats = 20)
NN$fitted
NN$residuals
NN$method#(x;y) 2 hidden layers each one has the value of neurons wrtten
NN$model # model information
summary(NN)

# Forecasting the NN
NNforecast=forecast(NN, h=25)
NNforecastdf=data.frame(NNforecast)
plot(forecastfit$forecastfit)
lines(NNforecastdf$Point.Forecast)

plot(NNforecastdf$Point.Forecast)
lines(forecastfit$forecastfit)
# Accuracy
NNaccuracy=accuracy(NNforecast)
NNaccuracydf=data.frame(NNaccuracy)
# Dataframe of the results
NNfitteddf=data.frame(NN$fitted)
NNresidualsdf=data.frame(NN$residuals)
NNmethoddf=data.frame(NN$method)
NN$model

NNmodeldf=data.frame("Average of 20 networks, each of which is
a 11-6-1 network with 79 weights options were - linear output units")
# Saving the results
write.xlsx(
  NNfitteddf,
  file="NN-dataset-.xlsx",
  sheetName = "Fitted Model",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE
)
write.xlsx(
  NNresidualsdf,
  file="NN-dataset-.xlsx",
  sheetName = "Residuals",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE
)
write.xlsx(
  NNmethoddf,
  file="NN-dataset-.xlsx",
  sheetName = "Method",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE
)
write.xlsx(
  NNmodeldf,
  file="NN-dataset-.xlsx",
  sheetName = "Model",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE
)
write.xlsx(
  NNforecastdf,
  file="NN-dataset-.xlsx",
  sheetName = "Forecast",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE
)
write.xlsx(
  NNaccuracydf,
  file="NN-dataset-.xlsx",
  sheetName = "Accuracy",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE
)




########################


#Determine hte order

acf(sarfimap)
pacf(sarfimap)
x=sarfimap$SARFIMA


acf(x)
pacf(x)

#Fitting the model
NN=nnetar(ts(x),p=11,P=1,repeats = 20)
NN$fitted
NN$residuals
NN$method#(x;y) 2 hidden layers each one has the value of neurons wrtten
NN$model # model information
summary(NN)

# Forecasting the NN
NNforecast=forecast(NN, h=25)
NNforecastdf=data.frame(NNforecast)
plot(forecastfit$forecastfit)
lines(NNforecastdf$Point.Forecast)

plot(NNforecastdf$Point.Forecast)
lines(forecastfit$forecastfit)
# Accuracy
NNaccuracy=accuracy(NNforecast)
NNaccuracydf=data.frame(NNaccuracy)
# Dataframe of the results
NNfitteddf=data.frame(NN$fitted)
NNresidualsdf=data.frame(NN$residuals)
NNmethoddf=data.frame(NN$method)
NN$model

NNmodeldf=data.frame("Average of 20 networks, each of which is
a 11-6-1 network with 79 weights options were - linear output units")
# Saving the results
write.xlsx(
  NNfitteddf,
  file="NN-sarfima-.xlsx",
  sheetName = "Fitted Model",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE
)
write.xlsx(
  NNresidualsdf,
  file="NN-sarfima-.xlsx",
  sheetName = "Residuals",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE
)
write.xlsx(
  NNmethoddf,
  file="NN-sarfima-.xlsx",
  sheetName = "Method",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE
)
write.xlsx(
  NNmodeldf,
  file="NN-sarfima-.xlsx",
  sheetName = "Model",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE
)
write.xlsx(
  NNforecastdf,
  file="NN-sarfima-.xlsx",
  sheetName = "Forecast",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE
)
write.xlsx(
  NNaccuracydf,
  file="NN-sarfima-.xlsx",
  sheetName = "Accuracy",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE
)

## Saving Results:
#Saving Input analysis
sd
summaryofdatabase.df= as.data.frame(sd)
summaryofdatabase.df
str(summaryofdatabase.df)

write.xlsx(
  sd,
  file="data analysis.xlsx",
  sheetName = "summary",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)

bpstats
bpstats.df <- as.data.frame(bpstats)
bpstats.df
str(bpstats.df)
write.xlsx(
  bpstats.df,
  file="data analysis.xlsx",
  sheetName = "boxplotstats",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

bpconf
bpconf.df <- as.data.frame(bpconf)
bpconf.df
str(bpconf.df)
write.xlsx(
  bpconf.df,
  file="data analysis.xlsx",
  sheetName = "boxplotconf",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

ag
ag.df<- as.data.frame(ag)
ag.df
str(ag.df)
write.xlsx(
  ag.df,
  file="data analysis.xlsx",
  sheetName = "Aggregate",
  col.names = TRUE,
  row.names = TRUE,
  append =TRUE,
  showNA= TRUE
  )

C02.AR
C02.AR.df <- as.data.frame(C02.AR)
C02.AR.df
str(C02.AR.df)
write.xlsx(
  C02.AR.df,
  file="data analysis.xlsx",
  sheetName = "Mean annual rate",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

months.ratio
months.ratio.df <- as.data.frame(months.ratio)
months.ratio.df
str(months.ratio.df)
write.xlsx(
  months.ratio.df,
  file="data analysis.xlsx",
  sheetName = "Month attribution",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

acfdatabase.df
str(acfdatabase.df)
write.xlsx(
  acfdatabase.df,
  file="data analysis.xlsx",
  sheetName = "AutoCorrelation for database",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

pacfdatabase.df
str(pacfdatabase.df)
write.xlsx(
  pacfdatabase.df,
  file="data analysis.xlsx",
  sheetName = "Partial AutoCorrelation for database",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

#Saving Decomposition

Decomposition$seasonal
Decompositionseasonal.df<- as.data.frame(Decomposition$seasonal)
Decompositionseasonal.df
str(Decompositionseasonal.df)

write.xlsx(
  Decompositionseasonal.df,
  file="Database Decomposition.xlsx",
  sheetName = "Seasonal component",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)

Decomposition$trend
Decomposition$trend.df <- as.data.frame(Decomposition$trend)
  Decomposition$trend.df
str(Decomposition$trend.df)
write.xlsx(
  Decomposition$trend.df,
  file="Database Decomposition.xlsx",
  sheetName = "Trend component",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

Decomposition$random
Decomposition$random.df <- as.data.frame(Decomposition$random)
Decomposition$random.df
str(Decomposition$random.df)
write.xlsx(
  Decomposition$random.df,
  file="Database Decomposition.xlsx",
  sheetName = "Random component",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

Decomposition$figure
Decomposition$figure.df <- as.data.frame(Decomposition$figure)
Decomposition$figure.df
str(Decomposition$figure.df)
write.xlsx(
  Decomposition$figure.df,
  file="Database Decomposition.xlsx",
  sheetName = "Figure for months ratio",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
decompositionmean
decompositionmean.df <- as.data.frame(decompositionmean)
decompositionmean.df
str(decompositionmean.df)
write.xlsx(
  decompositionmean.df,
  file="Database Decomposition.xlsx",
  sheetName = "Decomposition mean value",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

decompositionvar
decompositionvar.df <- as.data.frame(decompositionvar)
decompositionvar.df
str(decompositionvar.df)
write.xlsx(
  decompositionvar.df,
  file="Database Decomposition.xlsx",
  sheetName = "Decomposition variance value",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
error.Decomposition
error.Decomposition.df <- as.data.frame(error.Decomposition)
error.Decomposition.df 
str(error.Decomposition.df)
write.xlsx(
  error.Decomposition.df ,
  file="Database Decomposition.xlsx",
  sheetName = "Decomposition error",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
damagedcosine.df
str(damagedcosine.df)
write.xlsx(
  damagedcosine.df ,
  file="Database Decomposition.xlsx",
  sheetName = "AutoCorrelation of the random component",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
 

#####################################

 
#######################################

#Saving Tests
#Test for data analysis input

urkpsstestcriticalvalue.df<- as.data.frame(urkpsstestcriticalvalue)
urkpsstestcriticalvalue.df
str(urkpsstestcriticalvalue.df)
urkpssteststatistical.df<- as.data.frame(urkpssteststatistical)
urkpssteststatistical.df
str(urkpssteststatistical.df)
urkpsstestresiduals.df<- as.data.frame(urkpsstestresiduals)
urkpsstestresiduals.df
str(urkpsstestresiduals.df)
urkpsstestdecision.df<- as.data.frame(urkpsstestdecision)
urkpsstestdecision.df
str(urkpsstestdecision.df)

urkpsstestcriticalvalue.df
write.xlsx(
  urkpsstestcriticalvalue.df ,
  file="Database Analysis URKPSS Test.xlsx",
  sheetName = "Critical values",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)
urkpssteststatistical.df
write.xlsx(
  urkpssteststatistical.df ,
  file="Database Analysis URKPSS Test.xlsx",
  sheetName = "Statistical value",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
urkpsstestresiduals.df
write.xlsx(
  urkpsstestresiduals.df ,
  file="Database Analysis URKPSS Test.xlsx",
  sheetName = "Residuals values",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
urkpsstestdecision.df
write.xlsx(
  urkpsstestdecision.df ,
  file="Database Analysis URKPSS Test.xlsx",
  sheetName = "Decision result",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
acfurkpsstestresiduals.df
write.xlsx(
  acfurkpsstestresiduals.df ,
  file="Database Analysis URKPSS Test.xlsx",
  sheetName = "Autocorrelation residuals",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
pacfurkpsstestresiduals.df
write.xlsx(
  pacfurkpsstestresiduals.df ,
  file="Database Analysis URKPSS Test.xlsx",
  sheetName = "Partial Autocorrelation residuals",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

KpssTesttrendpvalue.df<- as.data.frame(KpssTesttrendpvalue)
KpssTesttrendpvalue.df
str(KpssTesttrendpvalue.df)
write.xlsx(
  KpssTesttrendpvalue.df ,
  file="Database Analysis KPSS Trend Test.xlsx",
  sheetName = "P value",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)

kpssTesttrendstatistic.df<- as.data.frame(kpssTesttrendstatistic)
kpssTesttrendstatistic.df
str(kpssTesttrendstatistic.df)
write.xlsx(
  kpssTesttrendstatistic.df ,
  file="Database Analysis KPSS Trend Test.xlsx",
  sheetName = "Statistical value",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
kpsstest.resulttrend.df<- as.data.frame(kpsstest.resulttrend)
kpsstest.resulttrend.df
str(kpsstest.resulttrend.df)
write.xlsx(
  kpsstest.resulttrend.df ,
  file="Database Analysis KPSS Trend Test.xlsx",
  sheetName = "Decision",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
KpssTestLevelpvalue
KpssTestLevelpvalue.df<- as.data.frame(KpssTestLevelpvalue)
KpssTestLevelpvalue.df
str(KpssTestLevelpvalue.df)
write.xlsx(
  KpssTestLevelpvalue.df ,
  file="Database Analysis KPSS Level Test.xlsx",
  sheetName = "P value",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)
kpssTestLevelstatistic
kpssTestLevelstatistic.df<- as.data.frame(kpssTestLevelstatistic)
kpssTestLevelstatistic.df
str(kpssTestLevelstatistic.df)
write.xlsx(
  kpssTestLevelstatistic.df ,
  file="Database Analysis KPSS Level Test.xlsx",
  sheetName = "Statistical value",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
kpsstest.resultLevel
kpsstest.resultLevel.df<- as.data.frame(kpsstest.resultLevel)
kpsstest.resultLevel.df
str(kpsstest.resultLevel.df)
write.xlsx(
  kpsstest.resultLevel.df ,
  file="Database Analysis KPSS Level Test.xlsx",
  sheetName = "Decision",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

adfpvalue
adfpvalue.df<- as.data.frame(adfpvalue)
adfpvalue.df
str(adfpvalue.df)
write.xlsx(
  adfpvalue.df ,
  file="Database Analysis ADF not a time trend Test.xlsx",
  sheetName = "P value",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)

adfcritical
adfcritical.df<- as.data.frame(adfcritical)
adfcritical.df
str(adfcritical.df)
write.xlsx(
  adfcritical.df ,
  file="Database Analysis ADF not a time trend Test.xlsx",
  sheetName = "Critical value",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
adfstatistic
adfstatistic.df<- as.data.frame(adfstatistic)
adfstatistic.df
str(adfstatistic.df)
write.xlsx(
  adfstatistic.df ,
  file="Database Analysis ADF not a time trend Test.xlsx",
  sheetName = "Statistical value",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
adfresult
adfresult.df<- as.data.frame(adfresult)
adfresult.df
str(adfresult.df)
write.xlsx(
  adfresult.df ,
  file="Database Analysis ADF not a time trend Test.xlsx",
  sheetName = "Decision",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

adfseaspvalue
adfseaspvalue.df<- as.data.frame(adfseaspvalue)
adfseaspvalue.df
str(adfseaspvalue.df)
write.xlsx(
  adfseaspvalue.df ,
  file="Database Analysis ADF intercept and a time trend Test.xlsx",
  sheetName = "P value",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)

adfseascritical
adfseascritical.df<- as.data.frame(adfseascritical)
adfseascritical.df
str(adfseascritical.df)
write.xlsx(
  adfseascritical.df ,
  file="Database Analysis ADF intercept and a time trend Test.xlsx",
  sheetName = "Critical value",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
adfseasstatistic
adfseasstatistic.df<- as.data.frame(adfseasstatistic)
adfseasstatistic.df
str(adfseasstatistic.df)
write.xlsx(
  adfseasstatistic.df ,
  file="Database Analysis ADF intercept and a time trend Test.xlsx",
  sheetName = "Statistical value",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
adfseasresult
adfseasresult.df<- as.data.frame(adfseasresult)
adfseasresult.df
str(adfseasresult.df)
write.xlsx(
  adfseasresult.df ,
  file="Database Analysis ADF intercept and a time trend Test.xlsx",
  sheetName = "Decision",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

# SARIMA seasonality
integrationseasonalvalue.df<- as.data.frame(integrationseasonalvalue)
integrationseasonalvalue.df
str(integrationseasonalvalue.df)
write.xlsx(
  integrationseasonalvalue.df ,
  file="SARIMA Integration value.xlsx",
  sheetName = "Integration seasonal value",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)

stationarywithseasonality
stationarywithseasonality.df<- as.data.frame(stationarywithseasonality)
stationarywithseasonality.df
str(stationarywithseasonality.df)
write.xlsx(
  stationarywithseasonality.df ,
  file="SARIMA Integration value.xlsx",
  sheetName = "Stationarity using Integration seasonal value",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

## Spectral density Function
#ARFIMA Spectral Desnsity
periodogramspectrum
periodogramspectrum.df<- as.data.frame(periodogramspectrum)
periodogramspectrum.df
str(periodogramspectrum.df)

write.xlsx(
  periodogramspectrum.df,
  file="Spectral Density ARFIMA Function.xlsx",
  sheetName = "Spectral Density",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)

periodogramfrequency
periodogramfrequency.df <- as.data.frame(periodogramfrequency)
periodogramfrequency.df 
str(periodogramfrequency.df)
write.xlsx(
  periodogramfrequency.df ,
  file="Spectral Density ARFIMA Function.xlsx",
  sheetName = "Frequency",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
period1
period1.df <- as.data.frame(period1)
period1.df 
str(period1.df)
write.xlsx(
  period1.df ,
  file="Spectral Density ARFIMA Function.xlsx",
  sheetName = "Period 1 Highest value",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
period2
period2.df <- as.data.frame(period2)
period2.df 
str(period1.df)
write.xlsx(
  period2.df ,
  file="Spectral Density ARFIMA Function.xlsx",
  sheetName = "Period 2 Highest value",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

#Raw Periodogram Spectral Desnsity
spectrumrawperiodogramfrequency
spectrumrawperiodogramfrequency.df<- as.data.frame(spectrumrawperiodogramfrequency)
spectrumrawperiodogramfrequency.df
str(spectrumrawperiodogramfrequency.df)

write.xlsx(
  spectrumrawperiodogramfrequency.df,
  file="Spectral Density RawPeriodogram Function.xlsx",
  sheetName = "Frequency",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)

spectrumrawperiodogramspectrum
spectrumrawperiodogramspectrum.df <- as.data.frame(spectrumrawperiodogramspectrum)
spectrumrawperiodogramspectrum.df 
str(spectrumrawperiodogramspectrum.df)
write.xlsx(
  spectrumrawperiodogramspectrum.df ,
  file="Spectral Density RawPeriodogram Function.xlsx",
  sheetName = "Spectral Density",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
spectrumrawperiodogramdegreeoffreedom
spectrumrawperiodogramdegreeoffreedom.df <- as.data.frame(spectrumrawperiodogramdegreeoffreedom)
spectrumrawperiodogramdegreeoffreedom.df 
str(spectrumrawperiodogramdegreeoffreedom.df)
write.xlsx(
  spectrumrawperiodogramdegreeoffreedom.df ,
  file="Spectral Density RawPeriodogram Function.xlsx",
  sheetName = "Degree of Freedom",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
spectrumrawperiodogrambandwidth
spectrumrawperiodogrambandwidth.df <- as.data.frame(spectrumrawperiodogrambandwidth)
spectrumrawperiodogrambandwidth.df 
str(spectrumrawperiodogrambandwidth.df)
write.xlsx(
  spectrumrawperiodogrambandwidth.df ,
  file="Spectral Density RawPeriodogram Function.xlsx",
  sheetName = "BandWindth",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
spectrumrawperiodogramtaper
spectrumrawperiodogramtaper.df <- as.data.frame(spectrumrawperiodogramtaper)
spectrumrawperiodogramtaper.df 
str(spectrumrawperiodogramtaper.df)
write.xlsx(
  spectrumrawperiodogramtaper.df ,
  file="Spectral Density RawPeriodogram Function.xlsx",
  sheetName = "Taper",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

#AutoRegression Periodogram Spectral Desnsity
spectrumARfrequency
spectrumARfrequency.df<- as.data.frame(spectrumARfrequency)
spectrumARfrequency.df
str(spectrumARfrequency.df)

write.xlsx(
  spectrumARfrequency.df,
  file="Spectral Density AutoRegression Function.xlsx",
  sheetName = "Frequency",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)

spectrumARspectrum
spectrumARspectrum.df <- as.data.frame(spectrumARspectrum)
spectrumARspectrum.df 
str(spectrumARspectrum.df)
write.xlsx(
  spectrumARspectrum.df ,
  file="Spectral Density AutoRegression Function.xlsx",
  sheetName = "Spectral Density",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
## Estimation of fractional integration value
#Using Periodogram
sperio.d
sperio.d.df<- as.data.frame(sperio.d)
sperio.d.df
str(sperio.d.df)

write.xlsx(
  sperio.d.df,
  file="Estimation of fd using periodogram.xlsx",
  sheetName = "Fractional integration value",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)

sperio.asymptotic.sd
sperio.asymptotic.sd.df <- as.data.frame(sperio.asymptotic.sd)
sperio.asymptotic.sd.df 
str(sperio.asymptotic.sd.df)
write.xlsx(
  sperio.asymptotic.sd.df ,
  file="Estimation of fd using periodogram.xlsx",
  sheetName = "Asymptotic standard deviation ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sperio.deviation.sd
sperio.deviation.sd.df <- as.data.frame(sperio.deviation.sd)
sperio.deviation.sd.df 
str(sperio.deviation.sd.df)
write.xlsx(
  sperio.deviation.sd.df ,
  file="Estimation of fd using periodogram.xlsx",
  sheetName = "Error standard deviation ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

#Using Gewek&porter-hudak
gph
gph.df<- as.data.frame(gph)
gph.df
str(gph.df)

write.xlsx(
  gph.df,
  file="Estimation of fd using GPH.xlsx",
  sheetName = "Fractional integration value",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)

#Using Multivariante local whittle estimator
GSE
GSE.df<- as.data.frame(GSE)
GSE.df
str(GSE.df)

write.xlsx(
  GSE.df,
  file="Estimation of fd using GSE.xlsx",
  sheetName = "Fractional integration value",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)
#Using modified local whittle estimator

HP.d
HP.d.df<- as.data.frame(HP.d)
HP.d.df
str(HP.d.df)

write.xlsx(
  HP.d.df,
  file="Estimation of fd using HP.xlsx",
  sheetName = "Fractional integration value",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)

HP.se
HP.se.df <- as.data.frame(HP.se)
HP.se.df 
str(HP.se.df)
write.xlsx(
  HP.se.df ,
  file="Estimation of fd using HP.xlsx",
  sheetName = "Standard Error ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
# Using local whittle estimator
LWnone.d
LWnone.d.df<- as.data.frame(LWnone.d)
LWnone.d.df
str(LWnone.d.df)

write.xlsx(
  LWnone.d.df,
  file="Estimation of fd using local whittle.xlsx",
  sheetName = "Robinson-Fractional integration value",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)

LWnone.se
LWnone.se.df <- as.data.frame(LWnone.se)
LWnone.se.df 
str(LWnone.se.df)
write.xlsx(
  LWnone.se.df ,
  file="Estimation of fd using local whittle.xlsx",
  sheetName = "Robinson-Standard Error ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
LWvelasco.d
LWvelasco.d.df <- as.data.frame(LWvelasco.d)
LWvelasco.d.df 
str(LWvelasco.d.df)
write.xlsx(
  LWvelasco.d.df ,
  file="Estimation of fd using local whittle.xlsx",
  sheetName = "Velasco-Fractional integration value ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
LWvelasco.se
LWvelasco.se.df <- as.data.frame(LWvelasco.se)
LWvelasco.se.df 
str(LWvelasco.se.df)
write.xlsx(
  LWvelasco.se.df ,
  file="Estimation of fd using local whittle.xlsx",
  sheetName = "Velasco-Standard Error ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
LWHC.d
LWHC.d.df <- as.data.frame(LWHC.d)
LWHC.d.df 
str(LWHC.d.df)
write.xlsx(
  LWHC.d.df ,
  file="Estimation of fd using local whittle.xlsx",
  sheetName = "Hurvich&Chen-Fractional integration value ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
LWHC.se
LWHC.se.df <- as.data.frame(LWHC.se)
LWHC.se.df 
str(LWHC.se.df)
write.xlsx(
  LWHC.se.df ,
  file="Estimation of fd using local whittle.xlsx",
  sheetName = "Hurvich&Chen-Standard Error ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

#Using local polynominal whittle estimator plus noise estimator

LPWNd
LPWNd.df<- as.data.frame(LPWNd)
LPWNd.df
str(LPWNd.df)

write.xlsx(
  LPWNd.df,
  file="Estimation of fd using polynomial local whittle & noise .xlsx",
  sheetName = "Frederiksem-Fractional integration value",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)

LPWNs.e
LPWNs.e.df <- as.data.frame(LPWNs.e)
LPWNs.e.df 
str(LPWNs.e.df)
write.xlsx(
  LPWNs.e.df ,
  file="Estimation of fd using polynomial local whittle & noise .xlsx",
  sheetName = "Frederiksem-Standard Error ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
#Estimation of the fractional integration value by range using the above methods
dwithoutvelsacoandHC
dwithoutvelsacoandHC.df<- as.data.frame(dwithoutvelsacoandHC)
dwithoutvelsacoandHC.df
str(dwithoutvelsacoandHC.df)

write.xlsx(
  dwithoutvelsacoandHC.df,
  file="Fractional Integration Range using 5 methods .xlsx",
  sheetName = "Velasco&HC not included-Fractional integration value",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)

dwithvelsacoandHC
dwithvelsacoandHC.df <- as.data.frame(dwithvelsacoandHC)
dwithvelsacoandHC.df 
str(dwithvelsacoandHC.df)
write.xlsx(
  dwithvelsacoandHC.df ,
  file="Fractional Integration Range using 5 methods .xlsx",
  sheetName = "Velasco&HC included-Fractional integration value ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

write.xlsx(
  acfusingfractionalintegration.df ,
  file="Fractional Integration Range using 5 methods .xlsx",
  sheetName = "AutoCorrelation using Fractional Integration ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
#hurst Exponent - Fractional Integration and d*
hurstexponent
write.xlsx(
  hurstexponent,
  file="HurstExponent-df-dfstar.xlsx",
  sheetName = "HurstExponent",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)
write.xlsx(
  df ,
  file="HurstExponent-df-dfstar.xlsx",
  sheetName = "Fd Simple RoverS hurst Fractional Integration",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
write.xlsx(
  dfcorrected ,
  file="HurstExponent-df-dfstar.xlsx",
  sheetName = "Fd corrected  RoverS hurst Fractional Integration",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
dstarARFIMA.df=as.data.frame(dstarARFIMA)
write.xlsx(
  dstarARFIMA.df ,
  file="HurstExponent-df-dfstar.xlsx",
  sheetName = "dstar simple  RoverS hurst Fractional Integration",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
dstarARFIMAcorrected.df=as.data.frame(dstarARFIMAcorrected)
write.xlsx(
  dstarARFIMAcorrected.df ,
  file="HurstExponent-df-dfstar.xlsx",
  sheetName = "dstar corrected  RoverS hurst Fractional Integration",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
#Qu Test for long memory process

QUstatistical
QUstatistical.df<- as.data.frame(QUstatistical)
QUstatistical.df
str(QUstatistical.df)

write.xlsx(
  QUstatistical.df,
  file="Qu Test for long memory process.xlsx",
  sheetName = "Statistical value",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)

QUcritical
QUcritical.df <- as.data.frame(QUcritical)
QUcritical.df 
str(QUcritical.df)
write.xlsx(
  QUcritical.df ,
  file="Qu Test for long memory process.xlsx",
  sheetName = "critical value ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
Qutestdecision
Qutestdecision.df <- as.data.frame(Qutestdecision)
Qutestdecision.df 
str(Qutestdecision.df)
write.xlsx(
  Qutestdecision.df ,
  file="Qu Test for long memory process.xlsx",
  sheetName = "Decision ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

#Fractional Cointegration Test

FCIN10statistic
FCIN10statistic.df<- as.data.frame(FCIN10statistic)
FCIN10statistic.df
str(FCIN10statistic.df)

write.xlsx(
  FCIN10statistic.df,
  file="Fractional Cointegration Test.xlsx",
  sheetName = "Statistical value",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)

FCIN10critical
FCIN10critical.df <- as.data.frame(FCIN10critical)
FCIN10critical.df 
str(FCIN10critical.df)
write.xlsx(
  FCIN10critical.df ,
  file="Fractional Cointegration Test.xlsx",
  sheetName = "critical value ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
FCIN10testdecisionbyreject
FCIN10testdecisionbyreject.df <- as.data.frame(FCIN10testdecisionbyreject)
FCIN10testdecisionbyreject.df 
str(FCIN10testdecisionbyreject.df)
write.xlsx(
  FCIN10testdecisionbyreject.df ,
  file="Fractional Cointegration Test.xlsx",
  sheetName = "Decision by reject ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
FCIN10testdecision
FCIN10testdecision.df <- as.data.frame(FCIN10testdecision)
FCIN10testdecision.df 
str(FCIN10testdecision.df)
write.xlsx(
  FCIN10testdecision.df ,
  file="Fractional Cointegration Test.xlsx",
  sheetName = "Decision",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

#SARFIMA MODEL PARAMETERS
sarfimaar
sarfimaar.df<- as.data.frame(sarfimaar)
sarfimaar.df
str(sarfimaar.df)

write.xlsx(
  sarfimaar.df,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "AR parameter",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)

sarfimama
sarfimama.df <- as.data.frame(sarfimama)
sarfimama.df 
str(sarfimama.df)
write.xlsx(
  sarfimama.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "MA parameter ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimasar
sarfimasar.df <- as.data.frame(sarfimasar)
sarfimasar.df 
str(sarfimasar.df)
write.xlsx(
  sarfimasar.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "Seasonal AR parameter ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimasma
sarfimasma.df <- as.data.frame(sarfimasma)
sarfimasma.df 
str(sarfimasma.df)
write.xlsx(
  sarfimasma.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "Seasonal MA parameter ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimad
sarfimad.df <- as.data.frame(sarfimad)
sarfimad.df 
str(sarfimad.df)
write.xlsx(
  sarfimad.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = " Integration parameter ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimasd
sarfimasd.df <- as.data.frame(sarfimasd)
sarfimasd.df 
str(sarfimasd.df)
write.xlsx(
  sarfimasd.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = " Seasonal Integration parameter ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimaH
sarfimaH.df <- as.data.frame(sarfimaH)
sarfimaH.df 
str(sarfimaH.df)
write.xlsx(
  sarfimaH.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = " H parameter ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimaHs
sarfimaHs.df <- as.data.frame(sarfimaHs)
sarfimaHs.df 
str(sarfimaHs.df)
write.xlsx(
  sarfimaHs.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = " Seasonal H parameter ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimaalpha
sarfimaalpha.df <- as.data.frame(sarfimaalpha)
sarfimaalpha.df 
str(sarfimaalpha.df)
write.xlsx(
  sarfimaalpha.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = " FDWN noise parameter ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimasalpha
sarfimasalpha.df <- as.data.frame(sarfimasalpha)
sarfimasalpha.df 
str(sarfimasalpha.df)
write.xlsx(
  sarfimasalpha.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "Seasonal FDWN noise parameter ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimadelta
sarfimadelta.df <- as.data.frame(sarfimadelta)
sarfimadelta.df 
str(sarfimadelta.df)
write.xlsx(
  sarfimadelta.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "Expected difference parameter ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimaomega
sarfimaomega.df <- as.data.frame(sarfimaomega)
sarfimaomega.df 
str(sarfimaomega.df)
write.xlsx(
  sarfimaomega.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "Variance Intercept parameter ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimafittedmean
sarfimafittedmean.df <- as.data.frame(sarfimafittedmean)
sarfimafittedmean.df 
str(sarfimafittedmean.df)
write.xlsx(
  sarfimafittedmean.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "mean value muHat parameter ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimaloglik
sarfimaloglik.df <- as.data.frame(sarfimaloglik)
sarfimaloglik.df 
str(sarfimaloglik.df)
write.xlsx(
  sarfimaloglik.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "log likelihood value ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimapars
sarfimapars.df <- as.data.frame(sarfimapars)
sarfimapars.df 
str(sarfimapars.df)
write.xlsx(
  sarfimapars.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "Optimization routine parameters ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimaparsp
sarfimaparsp.df <- as.data.frame(sarfimaparsp)
sarfimaparsp.df 
str(sarfimaparsp.df)
write.xlsx(
  sarfimaparsp.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "Optimization routine(P)parameters ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimase
sarfimase.df <- as.data.frame(sarfimase)
sarfimase.df 
str(sarfimase.df)
write.xlsx(
  sarfimase.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "Model Standard Errors ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimahess
sarfimahess.df <- as.data.frame(sarfimahess)
sarfimahess.df 
str(sarfimahess.df)
write.xlsx(
  sarfimahess.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "hessian routine parameters ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimafittedwithoutmean
sarfimafittedwithoutmean.df <- as.data.frame(sarfimafittedwithoutmean)
sarfimafittedwithoutmean.df 
str(sarfimafittedwithoutmean.df)
write.xlsx(
  sarfimafittedwithoutmean.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "fitted model without muHat(mean) ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimaresiduals
sarfimaresiduals.df <- as.data.frame(sarfimaresiduals)
sarfimaresiduals.df 
str(sarfimaresiduals.df)
write.xlsx(
  sarfimaresiduals.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "Model Residuals",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimasigma2
sarfimasigma2.df <- as.data.frame(sarfimasigma2)
sarfimasigma2.df 
str(sarfimasigma2.df)
write.xlsx(
  sarfimasigma2.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "Model Sigmasquare(Variance)",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimafitted.df
write.xlsx(
  sarfimafitted.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "Model Fitted with muHat(mean)",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
difforiginalwithfittedwithmean
difforiginalwithfittedwithmean.df <- as.data.frame(difforiginalwithfittedwithmean)
difforiginalwithfittedwithmean.df 
str(difforiginalwithfittedwithmean.df)
write.xlsx(
  difforiginalwithfittedwithmean.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "Diff. bet. database and Fitted Model",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
R2SARFIMA
R2SARFIMA.df <- as.data.frame(R2SARFIMA)
R2SARFIMA.df 
str(R2SARFIMA.df)
write.xlsx(
  R2SARFIMA.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "R2 using Observed+residuals",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
R2SARFIMAfittedwithmean
R2SARFIMAfittedwithmean.df <- as.data.frame(R2SARFIMAfittedwithmean)
R2SARFIMAfittedwithmean.df 
str(R2SARFIMAfittedwithmean.df)
write.xlsx(
  R2SARFIMAfittedwithmean.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "R2 using Model Fitted",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

write.xlsx(
  sarfimaacfresiduals ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "AutoCorrelation of Residuals of Model",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
write.xlsx(
  sarfimaacfrealfitted ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "AutoCorrelation of Fitted  Model",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
BT.Cstatistic
BT.Cstatistic.df <- as.data.frame(BT.Cstatistic)
BT.Cstatistic.df 
str(BT.Cstatistic.df)
write.xlsx(
  BT.Cstatistic.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "LJUNGBOX-Statistic value",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
BT.Cpvalue
BT.Cpvalue.df <- as.data.frame(BT.Cpvalue)
BT.Cpvalue.df 
str(BT.Cpvalue.df)
write.xlsx(
  BT.Cpvalue.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "LJUNGBOX-P value",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
BT.Cresult
BT.Cresult.df <- as.data.frame(BT.Cresult)
BT.Cresult.df 
str(BT.Cresult.df)
write.xlsx(
  BT.Cresult.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "LJUNGBOX-Decision",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
df.fit1ffitted.R2
df.fit1ffitted.R2.df <- as.data.frame(df.fit1ffitted.R2)
df.fit1ffitted.R2.df 
str(df.fit1ffitted.R2.df)
write.xlsx(
  df.fit1ffitted.R2.df ,
  file="SARFIMA Model Parameters.xlsx",
  sheetName = "SARFIMA Fitted model using residuals",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
#Prediction of the SARFIMA MODEL

sarfimapredict
sarfimapredict.df<- as.data.frame(sarfimapredict)
sarfimapredict.df
str(sarfimapredict.df)

write.xlsx(
  sarfimapredict.df,
  file="SARFIMA Prediction.xlsx",
  sheetName = "Prediction",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)

sarfimaexactsd
sarfimaexactsd.df <- as.data.frame(sarfimaexactsd)
sarfimaexactsd.df 
str(sarfimaexactsd.df)
write.xlsx(
  sarfimaexactsd.df ,
  file="SARFIMA Prediction.xlsx",
  sheetName = "Exact standard Error ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimalimitsd
sarfimalimitsd.df <- as.data.frame(sarfimalimitsd)
sarfimalimitsd.df 
str(sarfimalimitsd.df)
write.xlsx(
  sarfimalimitsd.df ,
  file="SARFIMA Prediction.xlsx",
  sheetName = "Limit standard Error ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimasigma2
sarfimasigma2.df <- as.data.frame(sarfimasigma2)
sarfimasigma2.df 
str(sarfimasigma2.df)
write.xlsx(
  sarfimasigma2.df ,
  file="SARFIMA Prediction.xlsx",
  sheetName = "Average Variance ",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimaexactvar
sarfimaexactvar.df <- as.data.frame(sarfimaexactvar)
sarfimaexactvar.df 
str(sarfimaexactvar.df)
write.xlsx(
  sarfimaexactvar.df ,
  file="SARFIMA Prediction.xlsx",
  sheetName = "Exact Variance Per step",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimalimitvar
sarfimalimitvar.df <- as.data.frame(sarfimalimitvar)
sarfimalimitvar.df 
str(sarfimalimitvar.df)
write.xlsx(
  sarfimalimitvar.df ,
  file="SARFIMA Prediction.xlsx",
  sheetName = "Limit Variance",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfimapredictconfidence
sarfimapredictconfidence.df <- as.data.frame(sarfimapredictconfidence)
sarfimapredictconfidence.df 
str(sarfimapredictconfidence.df)
write.xlsx(
  sarfimapredictconfidence.df ,
  file="SARFIMA Prediction.xlsx",
  sheetName = "Confidence 95percent",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

#Forecasting SARFIMA MODEL

df.forecast
df.forecast.df<- as.data.frame(df.forecast)
df.forecast.df
str(df.forecast.df)

write.xlsx(
  df.forecast.df,
  file="Forecasting SARFIMA Model.xlsx",
  sheetName = "Forecast all components",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)

sarfima.forecast
sarfima.forecast.df <- as.data.frame(sarfima.forecast)
sarfima.forecast.df 
str(sarfima.forecast.df)
write.xlsx(
  sarfima.forecast.df ,
  file="Forecasting SARFIMA Model.xlsx",
  sheetName = "Forecast component",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfima.Low.80
sarfima.Low.80.df <- as.data.frame(sarfima.Low.80)
sarfima.Low.80.df 
str(sarfima.Low.80.df)
write.xlsx(
  sarfima.Low.80.df ,
  file="Forecasting SARFIMA Model.xlsx",
  sheetName = "Low 80 CI component",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfima.Low.95
sarfima.Low.95.df <- as.data.frame(sarfima.Low.95)
sarfima.Low.95.df 
str(sarfima.Low.95.df)
write.xlsx(
  sarfima.Low.95.df ,
  file="Forecasting SARFIMA Model.xlsx",
  sheetName = "Low 95 CI component",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfima.High.95
sarfima.High.95.df <- as.data.frame(sarfima.High.95)
sarfima.High.95.df 
str(sarfima.High.95.df)
write.xlsx(
  sarfima.High.95.df ,
  file="Forecasting SARFIMA Model.xlsx",
  sheetName = "High 95 CI component",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sarfima.High.80
sarfima.High.80.df <- as.data.frame(sarfima.High.80)
sarfima.High.80.df 
str(sarfima.High.80.df)
write.xlsx(
  sarfima.High.80.df ,
  file="Forecasting SARFIMA Model.xlsx",
  sheetName = "High 80 CI component",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
forecastacfresid
forecastacfresid.df <- as.data.frame(forecastacfresid)
forecastacfresid.df 
str(forecastacfresid.df)
write.xlsx(
  forecastacfresid.df ,
  file="Forecasting SARFIMA Model.xlsx",
  sheetName = "ACF of residuals forecast fitted model",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
forecasterrorinfittedmodel
forecasterrorinfittedmodel.df <- as.data.frame(forecasterrorinfittedmodel)
forecasterrorinfittedmodel.df 
str(forecasterrorinfittedmodel.df)
write.xlsx(
  forecasterrorinfittedmodel.df ,
  file="Forecasting SARFIMA Model.xlsx",
  sheetName = "Error in forecast(database-forecastfitted)",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
forecastmodel
forecastmodel.df <- as.data.frame(forecastmodel)
forecastmodel.df 
str(forecastmodel.df)
write.xlsx(
  forecastmodel.df ,
  file="Forecasting SARFIMA Model.xlsx",
  sheetName = "Forecast parameters",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
diffinforecast
diffinforecast.df <- as.data.frame(diffinforecast)
diffinforecast.df 
str(diffinforecast.df)
write.xlsx(
  diffinforecast.df ,
  file="Forecasting SARFIMA Model.xlsx",
  sheetName = "Forecast-Predict",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
BT.Cfpvalue
BT.Cfpvalue.df <- as.data.frame(BT.Cfpvalue)
BT.Cfpvalue.df 
str(BT.Cfpvalue.df)
write.xlsx(
  BT.Cfpvalue.df ,
  file="Forecasting SARFIMA Model.xlsx",
  sheetName = "LJungBOX-P value",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
BT.Cfstatistic
BT.Cfstatistic.df <- as.data.frame(BT.Cfstatistic)
BT.Cfstatistic.df 
str(BT.Cfstatistic.df)
write.xlsx(
  BT.Cfstatistic.df ,
  file="Forecasting SARFIMA Model.xlsx",
  sheetName = "LJungBOX-Statistic value",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
BT.Cfresult
BT.Cfresult.df <- as.data.frame(BT.Cfresult)
BT.Cfresult.df 
str(BT.Cfresult.df)
write.xlsx(
  BT.Cfresult.df ,
  file="Forecasting SARFIMA Model.xlsx",
  sheetName = "LJungBOX-Decision",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

#Accuracy 

accuracysarfima
accuracysarfima.df<- as.data.frame(accuracysarfima)
accuracysarfima.df
str(accuracysarfima.df)

write.xlsx(
  accuracysarfima.df,
  file="Accuracy SARFIMA MODEL.xlsx",
  sheetName = "SARFIMA(muHat+fitted)",
  col.names = TRUE,
  row.names = TRUE,
  append = FALSE)

accuracysarfimaresidfitted
accuracysarfimaresidfitted.df <- as.data.frame(accuracysarfimaresidfitted)
accuracysarfimaresidfitted.df 
str(accuracysarfimaresidfitted.df)
write.xlsx(
  accuracysarfimaresidfitted.df ,
  file="Accuracy SARFIMA MODEL.xlsx",
  sheetName = "SARFIMA(Observed+residuals)",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
accuracyforecast
accuracyforecast.df <- as.data.frame(accuracyforecast)
accuracyforecast.df 
str(accuracyforecast.df)
write.xlsx(
  accuracyforecast.df ,
  file="Accuracy SARFIMA MODEL.xlsx",
  sheetName = "Forecast Model",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
absoluteerror
absoluteerror.df <- as.data.frame(absoluteerror)
absoluteerror.df 
str(absoluteerror.df)
write.xlsx(
  absoluteerror.df ,
  file="Accuracy SARFIMA MODEL.xlsx",
  sheetName = "Absolute Error",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
absolutepercenterror
absolutepercenterror.df <- as.data.frame(absolutepercenterror)
absolutepercenterror.df 
str(absolutepercenterror.df)
write.xlsx(
  absolutepercenterror.df ,
  file="Accuracy SARFIMA MODEL.xlsx",
  sheetName = "Absolute Percentage Error",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
averageactualbiggerthanpredicted
averageactualbiggerthanpredicted.df <- as.data.frame(averageactualbiggerthanpredicted)
averageactualbiggerthanpredicted.df 
str(averageactualbiggerthanpredicted.df)
write.xlsx(
  averageactualbiggerthanpredicted.df ,
  file="Accuracy SARFIMA MODEL.xlsx",
  sheetName = "averageactualbiggerthanpredicted",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
averageactualpercentbiggerthanpredicted
averageactualpercentbiggerthanpredicted.df <- as.data.frame(averageactualpercentbiggerthanpredicted)
averageactualpercentbiggerthanpredicted.df 
str(averageactualpercentbiggerthanpredicted.df)
write.xlsx(
  averageactualpercentbiggerthanpredicted.df ,
  file="Accuracy SARFIMA MODEL.xlsx",
  sheetName = "averageactualpercentbiggerthanpredicted",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
meanabsoluteerror
meanabsoluteerror.df <- as.data.frame(meanabsoluteerror)
meanabsoluteerror.df 
str(meanabsoluteerror.df)
write.xlsx(
  meanabsoluteerror.df ,
  file="Accuracy SARFIMA MODEL.xlsx",
  sheetName = "meanabsoluteerror",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
meanabsoluteprecenterror
meanabsoluteprecenterror.df <- as.data.frame(meanabsoluteprecenterror)
meanabsoluteprecenterror.df 
str(meanabsoluteprecenterror.df)
write.xlsx(
  meanabsoluteprecenterror.df ,
  file="Accuracy SARFIMA MODEL.xlsx",
  sheetName = "meanabsoluteprecenterror",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
meanabsolutescaleserror
meanabsolutescaleserror.df <- as.data.frame(meanabsolutescaleserror)
meanabsolutescaleserror.df 
str(meanabsolutescaleserror.df)
write.xlsx(
  meanabsolutescaleserror.df ,
  file="Accuracy SARFIMA MODEL.xlsx",
  sheetName = "meanabsolutescaleserror",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
medianabsoluteerror
medianabsoluteerror.df <- as.data.frame(medianabsoluteerror)
medianabsoluteerror.df 
str(medianabsoluteerror.df)
write.xlsx(
  medianabsoluteerror.df ,
  file="Accuracy SARFIMA MODEL.xlsx",
  sheetName = "medianabsoluteerror",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
meanquadraticweightedkappa
meanquadraticweightedkappa.df <- as.data.frame(meanquadraticweightedkappa)
meanquadraticweightedkappa.df 
str(meanquadraticweightedkappa.df)
write.xlsx(
  meanquadraticweightedkappa.df ,
  file="Accuracy SARFIMA MODEL.xlsx",
  sheetName = "meanquadraticweightedkappa",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
meansquarederror
meansquarederror.df <- as.data.frame(meansquarederror)
meansquarederror.df 
str(meansquarederror.df)
write.xlsx(
  meansquarederror.df ,
  file="Accuracy SARFIMA MODEL.xlsx",
  sheetName = "meansquarederror",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
meansquaredlogerror
meansquaredlogerror.df <- as.data.frame(meansquaredlogerror)
meansquaredlogerror.df 
str(meansquaredlogerror.df)
write.xlsx(
  meansquaredlogerror.df ,
  file="Accuracy SARFIMA MODEL.xlsx",
  sheetName = "meansquaredlogerror",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
relativeabsoluteerror
relativeabsoluteerror.df <- as.data.frame(relativeabsoluteerror)
relativeabsoluteerror.df 
str(relativeabsoluteerror.df)
write.xlsx(
  relativeabsoluteerror.df ,
  file="Accuracy SARFIMA MODEL.xlsx",
  sheetName = "relativeabsoluteerror",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
rootmeansquareerror
rootmeansquareerror.df <- as.data.frame(rootmeansquareerror)
rootmeansquareerror.df 
str(rootmeansquareerror.df)
write.xlsx(
  rootmeansquareerror.df ,
  file="Accuracy SARFIMA MODEL.xlsx",
  sheetName = "rootmeansquareerror",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
rootmeansquarelogerror
rootmeansquarelogerror.df <- as.data.frame(rootmeansquarelogerror)
rootmeansquarelogerror.df 
str(rootmeansquarelogerror.df)
write.xlsx(
  rootmeansquarelogerror.df ,
  file="Accuracy SARFIMA MODEL.xlsx",
  sheetName = "rootmeansquarelogerror",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
rootrelativesquarederror
rootrelativesquarederror.df <- as.data.frame(rootrelativesquarederror)
rootrelativesquarederror.df 
str(rootrelativesquarederror.df)
write.xlsx(
  rootrelativesquarederror.df ,
  file="Accuracy SARFIMA MODEL.xlsx",
  sheetName = "rootrelativesquarederror",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
relativesquarederror
relativesquarederror.df <- as.data.frame(relativesquarederror)
relativesquarederror.df 
str(relativesquarederror.df)
write.xlsx(
  relativesquarederror.df ,
  file="Accuracy SARFIMA MODEL.xlsx",
  sheetName = "relativesquarederror",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
squarederror
squarederror.df <- as.data.frame(squarederror)
squarederror.df 
str(squarederror.df)
write.xlsx(
  squarederror.df ,
  file="Accuracy SARFIMA MODEL.xlsx",
  sheetName = "squarederror",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
squaredlogerror
squaredlogerror.df <- as.data.frame(squaredlogerror)
squaredlogerror.df 
str(squaredlogerror.df)
write.xlsx(
  squaredlogerror.df ,
  file="Accuracy SARFIMA MODEL.xlsx",
  sheetName = "squaredlogerror",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
SymmetricMeanAbsolutePercentageError
SymmetricMeanAbsolutePercentageError.df <- as.data.frame(SymmetricMeanAbsolutePercentageError)
SymmetricMeanAbsolutePercentageError.df 
str(SymmetricMeanAbsolutePercentageError.df)
write.xlsx(
  SymmetricMeanAbsolutePercentageError.df ,
  file="Accuracy SARFIMA MODEL.xlsx",
  sheetName = "SymmetricMeanAbsolutePercentageError",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)
sumofsquarederrors
sumofsquarederrors.df <- as.data.frame(sumofsquarederrors)
sumofsquarederrors.df 
str(sumofsquarederrors.df)
write.xlsx(
  sumofsquarederrors.df ,
  file="Accuracy SARFIMA MODEL.xlsx",
  sheetName = "sumofsquarederrors",
  col.names = TRUE,
  row.names = TRUE,
  append = TRUE,
  showNA = TRUE
)

###########################################

###########



