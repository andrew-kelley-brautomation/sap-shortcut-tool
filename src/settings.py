import configparser, os

config = configparser.ConfigParser()

class SAP_Settings:
    #defaultRecordMailSubject = None
    #defaultRecordMailTime = None
    #defaultSolutionTime = None

    def __init__(self):
        #config.read_file(open(os.getcwd() + '\src\settings.ini'))
        #config.read_file(os.path.dirname(__file__) + '\settings.ini')
        config.read(os.path.join(os.path.dirname(__file__), 'settings.ini'))
        #DEFAULT
        self.defaultRecordMailSubject = config['DEFAULT']['DefaultRecordMailSubject']
        self.defaultRecordMailTime = config['DEFAULT']['DefaultRecordMailTime']
        self.defaultSolutionTime = config['DEFAULT']['DefaultSolutionTime']
        #Graphics
        self.graphicsScaling = config['Graphics']['GraphicsScaling']