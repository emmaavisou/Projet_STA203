library(Hmisc)
library(foreign)
library(epiDisplay)
library(epiR)
library(ggplot2)
library(readxl)
library(xlsx)


#############################################################
# Importation de la base
#############################################################
projet <- read_excel("Documents/Projet_STA203/SeoulBike.xlsx")
View(projet)
projet
str(projet)

#############################################################
# Renommer les colonnes
#############################################################
colnames(projet) <- c("Date","Count", "Hour", "Temp", "Hum", "Wind", "Visb", "Dew", "Solar", "Rain", "Snow", "Seasons", "Holiday", "Fday")

#############################################################
# Modifier le type des variables
#############################################################
projet$Date<- as.Date(projet$Date, format = "%d/%m/%Y")
projet$Seasons <- as.factor(projet$Seasons)
projet$Holiday <- as.factor(projet$Holiday)
projet$Fday <- as.factor(projet$Fday)
projet$Hour <- as.factor(projet$Hour)
projet$Temp <- as.numeric(projet$Temp)
projet$Wind <- as.numeric(projet$Wind)
projet$Dew <- as.numeric(projet$Dew)
projet$Solar <- as.numeric(projet$Solar)
projet$Rain <- as.numeric(projet$Rain)
projet$Snow <- as.numeric(projet$Snow)

str(projet)
describe(projet)

#############################################################
# Creation de variables pour l'analyse
#############################################################
projet$Raincat <- as.factor(ifelse(projet$Rain==0, 0, 
                         ifelse(projet$Rain<=1, 1,
                                ifelse(projet$Rain<=10, 2,3))))


#############################################################
# Analyse descriptive
#############################################################
tab1(projet$Seasons)
tab1(projet$Holiday)
tab1(projet$Fday)
tab1(projet$Raincat)

epi.descriptives(projet$Count)
epi.descriptives(projet$Temp)
epi.descriptives(projet$Hum)
epi.descriptives(projet$Wind)
epi.descriptives(projet$Visb)
epi.descriptives(projet$Dew)
epi.descriptives(projet$Solar)
# epi.descriptives(projet$Rain)
epi.descriptives(projet$Snow)
hist(projet$Snow)
hist(projet$Rain)

# Plot du nombre de vélo en fonction de l'heure
library(paletteer)
projet%>%
        ggplot(aes(Hour, Count, fill=Hour),color=stepped) +
        geom_smooth()+
        geom_boxplot(aes(group = Hour))+ theme(legend.position="none",panel.background = element_rect(fill = "white", colour = "black"))+xlab("Heures de la journée")+ylab("Nombre de vélos loués") 

#############################################################
# Matrice de correlation pour voir les variables qui sont liees
#############################################################
mat <- subset(projet, select=c("Count","Temp", "Hum", "Wind", "Visb", "Dew", "Solar", "Snow"))
M <- cor(mat)
library(corrplot)
corrplot(M, method = "number")

# symnum(M, abbr.colnames=FALSE)
# 
# library(corrplot)
# corrplot(M, type="upper", order="hclust", tl.col="black", tl.srt=45)

# On remarque que la temperature et la temperature au point de rose sont tres correles. On supprime une des deux valeurs du modele. 

# Mettre les variables Hour, pluie, neige, vent et soleil en quali pour voir (on cree de nouvelles variables, on ne remplace pas les initiales)
# projet$Hour1 <- as.factor(projet$Hour)
# projet$Neige <- ifelse(projet$Snow==0, "0", "1")
# projet$Pluie <- ifelse(projet$Rain==0, "0", "1")
# projet$Soleil <- ifelse(projet$Solar==0, "0", "1")
# projet$Vent <- ifelse(projet$Wind==0, "0", "1")

# hist(projet$Rain)
# str(projet)

#############################################################
# Analyse univariee des variables 
#############################################################
unipoi <- glm(Count~Snow, family=poisson(link = "log"), data = projet)
summary(unipoi)

unipoi1 <- glm(Count~Seasons, family=poisson(link = "log"), data = projet)
summary(unipoi1)
drop1(unipoi1,test="Chisq")

unipoi2 <- glm(Count~Holiday+Hour, family=poisson(link = "log"), data = projet)
summary(unipoi2)

unipoi3 <- glm(Count~Fday, family=poisson(link = "log"), data = projet)
summary(unipoi3)

unipoi4 <- glm(Count~Temp, family=poisson(link = "log"), data = projet)
summary(unipoi4)

unipoi5 <- glm(Count~Hum++Hour, family=poisson(link = "log"), data = projet)
summary(unipoi5)

unipoi6 <- glm(Count~Wind, family=poisson(link = "log"), data = projet)
summary(unipoi6)

unipoi7 <- glm(Count~Visb, family=poisson(link = "log"), data = projet)
summary(unipoi7)

unipoi8 <- glm(Count~Dew, family=poisson(link = "log"), data = projet)
summary(unipoi8)

unipoi9 <- glm(Count~Rain, family=poisson(link = "log"), data = projet)
summary(unipoi9)

unipoi10 <- glm(Count~Snow, family=poisson(link = "log"), data = projet)
summary(unipoi10)

unipoi10 <- glm(Count~Raincat, family=poisson(link = "log"), data = projet)
summary(unipoi10)
drop1(unipoi10,test="Chisq")


#############################################################
# Regression de poisson sans toucher aux variables, juste en enlevant la variable Dew
#############################################################
regpoi0 <- glm(Count~Hour+Temp+Hum+Wind+Visb+Solar+Rain+Snow+Holiday+Fday+Seasons, family=poisson(link = "log"), data = projet)
summary(regpoi0)

# AIC: 1066089

#############################################################
# Detection de potentielle surdispersion 
#############################################################
variance <- var(projet$Count)
esperance <- mean(projet$Count)
varespe <- variance/esperance #590.43
varespe
# Le rapport variance(Y)/esperance(Y) > 1, donc on peut supposer une surdispersion des observations par rapport à l'hypothèse de distribution de poisson. 

#############################################################
# Analyse de Poisson avec la variable Rain en 4 modalites, et sans la variable Dew
#############################################################
# Modele de poisson
regpoi1 <- glm(Count~Hour+Temp+Hum+Wind+Visb+Solar+Raincat+Snow+Holiday+Fday+Seasons, family=poisson(link = "log"), data = projet)
summary(regpoi1)

# Deviance = 4979261
# Deviance residuelle = 912667
# AIC: 979837
# summary(regpoi1)$dispersion

with(regpoi1, cbind(res.deviance = deviance, df = df.residual, rapport = deviance/df.residual))
# Deviance residuelle = 912666.9

with(regpoi1, cbind(pearson.chi2=sum(residuals(regpoi1, type = "pearson")^2), df = df.residual, rapport = sum(residuals(regpoi1, type = "pearson")^2)/df.residual))

# Chi 2 de Pearson = 881334.1
# Permet aussi d'avoir le Chi2 de Parson : sum(((projet$Count-regpoi1$fitted.values)^2)/regpoi1$fitted.values)
# Ou encore une autre méthode : sum(residuals(regpoi1,type="pearson")^2)
# Pour savoir si le modele de Poisson est adéquat, on calcul le ratio chi2 de Pearson / DDL. si ce n'est pas proche de 1 alors ce modele n'est pas adéquat.
# On a donc : 881334.1 / 8722 = 101.0473 --> Tres loin de 1 donc pas adéquat.
# Le ratio deviance / DDL est aussi une possibilite pour savoir l'adequation du modele 
# Mais mieux de prendre le Chi2 de Pearson

round(cbind(exp(coef(regpoi1)),exp(confint.default(regpoi1))),6)

#############################################################
# Calcul de la statistique Ta -> Test de surdispersion propose par Dean
#############################################################
Y_pred <- predict(regpoi1, type = "response")
projet$predpois <-Y_pred
library(boot)
diag_poisson <- glm.diag(regpoi1)
Ta <- (sum((projet$Count-Y_pred)^2 - projet$Count+glm.diag(regpoi1)$h*Y_pred))/((2*sum(Y_pred^2))^0.5)
Ta

# De plus, si le test n'est pas significatif, mais que la deviance du modele (ou la somme des residus de Pearson) est superieur au nombre de degres de libertes, alors on conclu egalement qu'il y a surdispersion


library(AER)
y <- dispersiontest(regpoi1)
# Il faut tester le modele de Quasi-Poisson et le modele Binomal negatif pour savoir quel modele est le plus adequat


#############################################################
# Quasi-poison
#############################################################
Quasipoi <-glm(Count~Hour+Temp+Hum+Wind+Visb+Solar+Raincat+Snow+Holiday+Fday+Seasons,  family = quasipoisson(link = log),data=projet)
summary(Quasipoi)

# On enleve la variable la moins significative : Fday
Quasipoi2 <-glm(Count~Hour+Temp+Hum+Wind+Visb+Solar+Raincat+Snow+Holiday+Seasons,  family = quasipoisson(link = log),data=projet)
summary(Quasipoi2)

# On enleve la variable la moins significative : Visb
Quasipoi3 <-glm(Count~factor(Hour)+Temp+Hum+Wind+Solar+factor(Raincat)+Snow+factor(Holiday)+factor(Seasons), family = quasipoisson(link = log),data=projet)
summary(Quasipoi3)

# On a besoin de la p-value globale pour plusieurs variable
drop1(Quasipoi3,tes="Chisq")

# Calcul des RR et de IC95%
exp(Quasipoi3$coefficients)
exp(confint.default(Quasipoi3))
confint(Quasipoi3)

#############################################################
# Binomiale négative
#############################################################
library(MASS)
Binneg <-glm.nb(Count~Hour+Temp+Hum+Wind+Visb+Solar+Raincat+Snow+Holiday+Fday+Seasons,data=projet)
summary(Binneg)
# AIC binomial negative = 115 489
# Modele de poisson avec les meme variables que le modele binomiale negative : AIC = 979 837
# Meilleur modele 
# Le modele binomial est donc preferable au modele de poisson 

# On enleve la variable la moins significative : Fday
Binneg2 <-glm.nb(Count~Hour+Temp+Hum+Wind+Visb+Solar+Raincat+Snow+Holiday+Seasons,data=projet)
summary(Binneg2)

# On enleve la variable la moins significative : Solar
Binneg3 <-glm.nb(Count~Hour+Temp+Hum+Wind+Visb+Raincat+Snow+Holiday+Seasons,data=projet)
summary(Binneg3)

# On enleve la variable la moins significative : Visb
Binneg4 <-glm.nb(Count~factor(Hour)+Temp+Hum+Wind+factor(Raincat)+Snow+factor(Holiday)+factor(Seasons),data=projet)
summary(Binneg4)
exp(Binneg4)
# Le modele binomial est donc preferable au modele de poisson 


# Il faut maintenant choisir entre le modele de quasi poisson et le modele binomial celui qui est le plus adequat

#############################################################
# Comparaison binomial et quasi-poisson
#############################################################
bin <-glm.nb(Count~factor(Hour)+Temp+Hum+Wind+factor(Raincat)+Snow+factor(Holiday)+factor(Seasons)+Solar+Visb+factor(Fday),data=projet)
bin
Quasipoi<-glm(Count~factor(Hour)+Temp+Hum+Wind+Solar+factor(Raincat)+Snow+factor(Holiday)+factor(Seasons)+Visb+factor(Fday), family = quasipoisson(link = log),data=projet)
regpoi

projet$predcomplet <- bin$fitted.values
projet$predquasicomplet <- Quasipoi$fitted.values

regpoi <-glm(Count~factor(Hour)+Temp+Hum+Wind+Solar+factor(Raincat)+Snow+factor(Holiday)+factor(Seasons)+Visb+factor(Fday), family = poisson(link = log),data=projet)
regpoi

xb <- predict(bin)#binomiale négative
g <- cut(xb, breaks=quantile(xb,seq(0,100,5)/100))
m <- tapply(projet$Count, g, mean)
v <- tapply(projet$Count, g, var)
pr <- residuals(regpoi,"pearson")
phi <- sum(pr^2)/df.residual(regpoi)
phi
x <- seq(2,2000,10)

ggplot()+aes(m,v)+geom_point(size=2)+ geom_smooth(aes(color="#99CC33"), se=F)+geom_line(aes(x, x*phi,color="blue"))+geom_line(aes(x, x*(1+x/bin$theta),color="red"))+
        scale_color_identity(name = "",
                             breaks = c("red", "#99CC33", "blue"),
                             labels = c("Binomiale-négative", "Lissage points (référence)", "Quasi-Poisson"),
                             guide = "legend")+
        theme_classic()+xlab("Moyenne")+ylab("Variance") +ylim(0,400000) + theme(legend.position = "top")

#############################################################
# Calcul du RMSE entre bin et quasi poisson complet
#############################################################
sqrt(mean((projet$Count - projet$predcomplet)^2))
sqrt(mean((projet$Count - projet$predquasicomplet)^2))


#############################################################
# Différence moyenne entre valeurs prédites par Quasi-Poisson 
#  et valeurs réelles
#############################################################
diff_quasi<- abs(projet$Count - fitted(Quasipoi3))
mean(diff_quasi)

#############################################################
# Graphe de la prédiction des données par rapport aux données 
# réelles
#############################################################
df <- subset(projet, Date=="2018-11-30")
projet$pred <- Quasipoi3$fitted.values
ggplot(df)+
        geom_line(aes(Hour,Count,color="black"))+
        geom_line(aes(Hour,pred, color="blue"))+
        scale_color_identity(name = "",
                             breaks = c("black", "blue"),
                             labels = c("Valeurs réelles (référence)", "Valeurs prédites"),
                             guide = "legend")+
        theme_classic()+xlab("Heure de la journée")+ylab("Nombre de vélos loués")+ theme(legend.position = "top")