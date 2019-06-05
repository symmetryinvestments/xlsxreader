handWrittenFiles=

style_lint:
	#@echo "Enforce braces on the same line"
	#grep -nrE '^[\t ]*{' $$(find source -name '*.d'); test $$? -eq 1

	@echo "Check for whitespace indentation"
	grep -nr '^[ ]' source ; test $$? -eq 1

	@echo "Enforce whitespace before opening parenthesis"
	grep -nrE "\<(for|foreach|foreach_reverse|if|while|switch|catch|version) \(" $$(find source -name '*.d') ; test $$? -eq 1

	@echo "Enforce no whitespace after opening parenthesis"
	grep -nrE "\<(version) \( " $$(find source -name '*.d') ; test $$? -eq 1

	@echo "Enforce whitespace between colon(:) for import statements (doesn't catch everything)"
	grep -nr 'import [^/,=]*:.*;' $$(find source -name '*.d') | grep -vE "import ([^ ]+) :\s"; test $$? -eq 1

	@echo "Check for package wide std.algorithm imports"
	grep -nr 'import std.algorithm : ' $$(find source -name '*.d') ; test $$? -eq 1

	@echo "Enforce no space between assert and the opening brace, i.e. assert("
	grep -nrE 'assert \(' $$(find source -name '*.d') ; test $$? -eq 1

	@echo "Enforce space between a .. b"
	grep -nrE '[[:alnum:]][.][.][[:alnum:]]|[[:alnum:]] [.][.][[:alnum:]]|[[:alnum:]][.][.] [[:alnum:]]' $$(find source -name '*.d' | grep -vE 'std/string.d|std/uni.d') ; test $$? -eq 1

	@echo "Enforce space between binary operators"
	grep -nrE "[[:alnum:]](==|!=|<=|<<|>>|>>>|^^)[[:alnum:]]|[[:alnum:]] (==|!=|<=|<<|>>|>>>|^^)[[:alnum:]]|[[:alnum:]](==|!=|<=|<<|>>|>>>|^^) [[:alnum:]]" $$(find source -name '*.d'); test $$? -eq 1

	@echo "Check line length"
	grep -nr '.\{81\}' $$(find source -name '*.d') ; test $$? -eq 1
