doc_file = odsfile.pdf
luafiles = odsfile.lua
content = odsfile.sty odsfile.tex odsfile.pdf odsfile.lua README pokus.ods
buildfolder = build/odsfile

doc: $(doc_file) odsfile.tex
	lualatex odsfile

build: doc
	@rm -rf build
	@mkdir -p $(buildfolder)
	@cp $(content) $(buildfolder)
	cd build && zip -r odsfile.zip odsfile

