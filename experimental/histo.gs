binwidth=0.5
set boxwidth binwidth
# set yrange [0:100]
# set xrange [0:2200]
bin(x,width)=width*floor(x/width) + binwidth/2.0
plot 'imp.dat' using (bin($1,binwidth)):(1.0) smooth freq with boxes
