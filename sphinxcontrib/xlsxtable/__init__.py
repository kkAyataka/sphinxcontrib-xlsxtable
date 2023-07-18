try:
    from . import xlsxtable
except ModuleNotFoundError:
    from sphinxcontrib.xlsxtable import xlsxtable

    if __name__ != "sphinxcontrib.xlsxtable":
        raise


def setup(app):
    xlsxtable.setup(app)
