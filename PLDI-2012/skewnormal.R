library("VGAM")
library("ggplot2")
samples = 10000000

skewnormal = rsnorm(n=samples,location=0,shape=-10,scale=1 )
skewnormal <- (skewnormal - mean(skewnormal)) / sd(skewnormal)

normal = rsnorm(n=samples,location=0,shape=0,scale=1 )
normal <- (normal - mean(normal)) / sd(normal)

dist <- c(rep('Skewed', length(skewnormal)), rep('Normal', length(normal)))
data <- data.frame(dist=dist, x=c(skewnormal, normal))
qplot(x=x, data=data, color=dist, stat='density', geom='line') + geom_vline(x=c(2, -2), linetype="dotted") + theme(legend.title=element_blank()) + xlab("X") + ylab("Density") + ggtitle("Standard Normal vs. Skewed Normal Distribution with 95% Confidence Interval")

