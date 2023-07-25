try:
    from . import xlsxtable
except ModuleNotFoundError:
    if __name__ != 'sphinxcontrib.xlsxtable':
        raise

def setup(app):
    xlsxtable.setup(app)
