MAIN = project

PDFLATEX = pdflatex -halt-on-error
BIBTEX = bibtex
PATCH = patch -p0 -t

R ?= HEAD
DIFF_DIR = .diff.$(R)
OLDVERSION_DIR = .oldversion.$(R)
CHANGES_FILE = .changes.$(R).diff

SIDE_EFFECTS = $(MAIN).pdf $(MAIN).out $(MAIN).aux $(MAIN).bbl $(MAIN).blg $(MAIN).log diff.pdf \
	$(wildcard diff.*) $(wildcard oldversion.*) $(wildcard changes.*)

FILES = $(filter-out $(SIDE_EFFECTS), $(wildcard *))
TEX_FILES = $(wildcard *.tex) $(wildcard berger/*.tex) $(wildcard kaplan/*.tex) $(wildcard guha/*.tex) $(wildcard richards/*.tex)
NONTEX_FILES = $(filter-out $(TEX_FILES), $(FILES))

OLD_FILES = $(addprefix $(OLDVERSION_DIR)/, $(FILES))
OLD_TEX_FILES = $(addprefix $(OLDVERSION_DIR/, $(TEX_FILES))
OLD_NONTEX_FILES = $(addprefix $(OLDVERSION_DIR/, $(NONTEX_FILES))

DIFF_FILES = $(addprefix $(DIFF_DIR)/, $(FILES))
DIFF_TEX_FILES = $(addprefix $(DIFF_DIR)/, $(TEX_FILES))
DIFF_NONTEX_FILES = $(addprefix $(DIFF_DIR)/, $(NONTEX_FILES))

all: $(MAIN).pdf $(MAIN).bbl

clean:
	@rm -f $(MAIN).pdf $(MAIN).out $(MAIN).aux $(MAIN).bbl $(MAIN).blg $(MAIN).log
	@rm -rf diff.pdf .oldversion.* .diff.* .changes.*

$(MAIN).pdf: $(wildcard *.tex)
	@echo "[MAIN] Running pdflatex"
	@$(PDFLATEX) $(MAIN).tex | tail

$(MAIN).bbl: $(wildcard *.bib)
	@echo "[MAIN] Running bibtex"
	@$(BIBTEX) $(MAIN).aux

	@echo "[MAIN] Updating document references"
	@$(PDFLATEX) $(MAIN).tex > /dev/null
	@$(PDFLATEX) $(MAIN).tex > /dev/null

diff: diff.pdf

diff.pdf: $(DIFF_DIR)/$(MAIN).pdf $(DIFF_DIR)/$(MAIN).bbl
	@cp $(DIFF_DIR)/$(MAIN).pdf diff.pdf

$(DIFF_DIR)/$(MAIN).pdf: $(DIFF_NONTEX_FILES) $(DIFF_DIR)/$(MAIN).tex
	@echo "[DIFF] Running pdflatex"
	@cd $(DIFF_DIR); $(PDFLATEX) $(MAIN).tex | tail

$(DIFF_DIR)/$(MAIN).bbl: $(wildcard $(DIFF_DIR)/*.bib)
	@echo "[DIFF] Running bibtex"
	@cd $(DIFF_DIR); $(BIBTEX) $(MAIN).aux
	@echo "[DIFF] Updating document references"
	@cd $(DIFF_DIR); $(PDFLATEX) $(MAIN).tex > /dev/null
	@cd $(DIFF_DIR); $(PDFLATEX) $(MAIN).tex > /dev/null

$(DIFF_DIR)/$(MAIN).tex: oldversion_dummy
	@echo "[DIFF] Running latexdiff"
	@mkdir -p $(DIFF_DIR)
	@latexdiff --flatten $(OLDVERSION_DIR)/$(MAIN).tex $(MAIN).tex > $@ 2> /dev/null

$(DIFF_NONTEX_FILES): oldversion_dummy
	@echo "[DIFF] Copying $(notdir $@) to diff workspace"
	@mkdir -p $(DIFF_DIR)
	@cp -R $(notdir $@) $@

oldversion_dummy: $(CHANGES_FILE)
	@echo "[DIFF] Copying files to old version workspace"
	@mkdir -p $(OLDVERSION_DIR)
	@cp -R $(FILES) $(OLDVERSION_DIR)/
	@echo "[DIFF] Applying changes to old version workspace"
	@$(PATCH) -R -d $(OLDVERSION_DIR) -i ../$(CHANGES_FILE)

$(CHANGES_FILE): $(FILES)
ifeq ($(R),HEAD)
	@echo "[DIFF] Using local checkout to compute changes"
	@svn diff *.tex > $@
else
	@echo "[DIFF] Fetching changes from repository"
	@svn diff -r $(R):HEAD > $@
endif

