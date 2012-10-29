all:	main.pdf

main.pdf:	main.tex abstract.tex introduction.tex background.tex future.tex algorithm.tex conclusion.tex evaluation.tex overview.tex related.tex ref.bib 
	pdflatex main.tex
	pdflatex main.tex

clean:
	rm -f main.log main.aux main.pdf

