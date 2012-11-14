library("VGAM")
library("ggplot2")

rwigner <- function(n, r) {
    result = rep(NA, n)
    i = 1
    while(i <= n) {
        x = runif(1, -r, r)[1]
        y = sqrt(r^2-x^2)
        p = runif(1)[1]
        if(p <= y) {
            result[i] = x
            i = i + 1
        }
    }
    
    return(result)
}

samples = 1000000

skewnormal = rsnorm(n=samples,location=0,shape=-10,scale=1 )
skewnormal <- (skewnormal - mean(skewnormal)) / sd(skewnormal)

normal = rsnorm(n=samples,location=0,shape=0,scale=1 )
normal <- (normal - mean(normal)) / sd(normal)

wigner = rwigner(n=samples, r=1)
wigner <- (wigner - mean(wigner)) / sd(wigner)

dist <- c(rep('Skewed', length(skewnormal)), rep('Normal', length(normal)), rep('Wigner', length(wigner)))
data <- data.frame(dist=dist, x=c(skewnormal, normal, wigner))
qplot(x=x, data=data, color=dist, stat='density', geom='line') + geom_vline(x=c(2, -2), linetype="dotted") + theme(legend.title=element_blank()) + xlab("X") + ylab("Density") + ggtitle("Standard Normal vs. Skewed Normal Distribution with 95% Confidence Interval")

