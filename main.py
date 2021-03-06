import glob
import os

from utils import utils
from utils import ParseProfile
from utils import ParsePermissionset
from utils import ParseProfileNew

def main():
    configObj = utils.getConfigInfo('config.yaml')

    if configObj['ParserPermissionSet'] == 1:
        print('### ParserPermissionSet start ###')
        ParsePermissionset.outputXmlDataToExcel()
        print('### ParserPermissionSet end ###')

    if configObj['ParserProfile'] == 1:
        print('### ParserProfile start ###')
        ParseProfileNew.outputXmlDataToExcel()
        print('### ParserProfile end ###')

    if configObj['parserProfileNote'] == 1:
        print('### parserProfileNote start ###')
        ParseProfile.parserProfileNote(configObj, True)
        print('### parserProfileNote end ###')

    if configObj['outputFileToCsvByList'] == 1:
        print('### outputXmlDataToCsvByList start ###')
        ParseProfile.outputXmlDataToCsvByList(configObj, True)
        print('### outputXmlDataToCsvByList end ###')

    if configObj['OutputFileToCsvByMatrix'] == 1:
        print('### OutputFileToCsvByMatrix start ###')
        ParseProfile.outputXmlDataToCsvByMatrix(configObj, True)
        print('### OutputFileToCsvByMatrix end ###')

    if configObj['OutputFileToExlByMatrix'] == 1:
        print('### OutputFileToExlByMatrix start ###')
        ParseProfile.outputXmlDataToExcelByMatrix(configObj, True)
        print('### OutputFileToExlByMatrix end ###')

if __name__ == "__main__":
    # calling main function
    main()