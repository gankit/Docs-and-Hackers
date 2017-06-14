#!/usr/bin/env python

# make sure to install these packages before running:
# pip install pandas
# pip install bokeh

import numpy as np
import pandas as pd
import datetime
import urllib
 
from bokeh.plotting import *
from bokeh.models import HoverTool
from collections import OrderedDict

query = ("https://openpaymentsdata.cms.gov/resource/t3za-xhk7.json?$$app_token=Q15zNm0GJ6aJeNjpKoAqqtE10")
raw_data = pd.read_json(query)
print(raw_data)