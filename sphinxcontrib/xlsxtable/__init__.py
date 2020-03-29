try:
    from . import xlsxtable
except:
    if __name__ != 'sphinxcontrib.xlsxtable':
        raise

def setup(app):
    xlsxtable.setup(app)
