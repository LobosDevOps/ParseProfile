import glob
import os

from utils import utils
from utils import ParseProfile

def main():
    configObj = utils.getConfigInfo('config.yaml')

    if configObj['parseprofile'] == 1:
        print('### outputXmlDataToCsvByList start ###')
        ParseProfile.outputXmlDataToCsvByList(configObj, True)
        print('### outputXmlDataToCsvByList end ###')

    if configObj['parseprofileToMatrix'] == 1:
        print('### outputXmlDataToCsvByMatrix start ###')
        ParseProfile.outputXmlDataToCsvByMatrix(configObj, True)
        # ParseProfile.outputXmlDataToExcelByMatrix(configObj, True)
        print('### outputXmlDataToCsvByMatrix end ###')

if __name__ == "__main__":
    # calling main function
    main()