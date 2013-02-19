all: install_deps test

filename=excellent-`python -c 'import excellent;print excellent.version'`.tar.gz

localshop="http://localshop.staging.yipit.com:8900/"

install_deps:
	@pip install -r requirements.txt

unit: clean
	@nosetests --with-coverage --stop --cover-package=excellent --verbosity=2 -s tests/unit/

functional: clean
	@nosetests --with-coverage --stop --cover-package=excellent --verbosity=2 -s tests/functional/

docs:
	@steadymark README.md

test: unit functional docs

clean:
	@printf "Cleaning up files that are already in .gitignore... "
	@for pattern in `cat .gitignore`; do rm -rf $$pattern; find . -name "$$pattern" -exec rm -rf {} \;; done
	@echo "OK!"
	@rm -f .coverage

release: clean test publish
	@printf "Exporting to $(filename)... "
	@tar czf $(filename) excellent setup.py README.md
	@echo "DONE!"

publish:
	@python setup.py register -r localshop
	@python setup.py sdist upload -r localshop
