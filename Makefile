CUSTOM_PIP_INDEX=public

all: install_deps test

filename=excellent-`python -c 'import excellent;print excellent.version'`.tar.gz

export PYTHONPATH:= ${PWD}

install_deps:
	@pip install -r requirements.txt

unit: clean
	@nosetests --with-coverage --stop --cover-package=excellent --verbosity=2 -s tests/unit/

functional: clean
	@nosetests --with-coverage --stop --cover-package=excellent --verbosity=2 -s tests/functional/

documentation:
	@steadymark README.md

test: unit functional documentation

clean:
	@printf "Cleaning up files that are already in .gitignore... "
	@for pattern in `cat .gitignore`; do rm -rf $$pattern; find . -name "$$pattern" -exec rm -rf {} \;; done
	@echo "OK!"
	@rm -f .coverage

release: clean test publish
	@printf "Exporting to $(filename)... "
	@tar czf $(filename) excellent setup.py README.md COPYING
	@echo "DONE!"

publish:	
	echo "Uploading to '$(CUSTOM_PIP_INDEX)'"
	python setup.py sdist upload -r "$(CUSTOM_PIP_INDEX)"
