
all: monad_notes

monad_notes:
	pandoc --from=markdown+pandoc_title_block --to=html --metadata pagetitle="Monad Notes" --standalone --output=output/monad_notes.html monad_notes.md
	pandoc --from=markdown --pdf-engine=pdflatex --standalone --output=output/monad_notes.pdf monad_notes.md
	pandoc --from=markdown --to=docx --reference-doc=include/custom-reference1.docx --standalone --output=output/monad_notes.docx monad_notes.md
