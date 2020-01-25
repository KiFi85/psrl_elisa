import pandas as pd
import numpy as np
import matplotlib as mpl
import matplotlib.dates as mpdt
import datetime as dt
import matplotlib.pyplot as plt
from pylab import arange, yticks
import calendar
from matplotlib.dates import MONTHLY, DateFormatter, rrulewrapper, RRuleLocator
import math

# Return a blue colour map
def getColourMap(colour_map):
    """ Returns sample colour map """

    # Saved colour map
    colMapdf = pd.read_csv(colour_map, header=None)
    # Convert to numpy array (depends on version of pandas)
    #colMap = colMapdf.to_numpy()
    colMap = colMapdf.values
    # Normalise between (0,1)    
    colMap = colMap/255
    return colMap

def getPlotDates(df):
    """ return the start and end dates for plotting """
    # Get start and end dates as floats
    # First convert to datetime
    df['start'] = pd.to_datetime(df['start'])
    df['end'] = pd.to_datetime(df['end'])
    
    start = df['start'].tolist()
    
    # Convert end and startdates to floats
    eDates, sDates = [mpl.dates.date2num(item) for item in 
                      (df['end'].tolist(), df['start'].tolist())]
    
    # Return dates
    return(sDates, eDates)
    
def getDeadlines(df,ydict):
    """ Get the dates for plotting shipments """
    
    # Convert dates to datetime
    # Drop rows without deadline
    df = df.dropna()
    subDates = pd.to_datetime(df['deadline'].tolist())
    
    # Get dates as float
    subDates = mpl.dates.date2num(subDates)
    
    # Get positions
    ylabels = df['sponsor'].map(str) + '\n' + df['study'].map(str)    
    ylabels = ylabels.tolist()
    
    ypos = [ydict[x] for x in ylabels]
    
    return (subDates,ypos)
        
def getYAxisDetails(df):
    """ Get the ylabels and y-positions for plotting """    
    
    # Get all possible ylabels including duplicates as list
    ylabels = df['sponsor'].map(str) + '\n' + df['study'].map(str)
        
    ylabels = ylabels.tolist()
    
    # Get unique values from list to deal with multiple instances/shipments
#    ylabelsUnique = list(set(ylabels))    
    seen = set()
    ylabelsUnique = [x for x in ylabels if not (x in seen or seen.add(x))]
    
    # Determine the top position and the possible range of values
    # Each position separated by 0.5
    topY = 0.5*len(ylabelsUnique)+0.5
    ypos = arange(0.5,topY,0.5)
    
    # Create a dictionary of unique values and their positions
    ydict = dict(zip(ylabelsUnique,ypos))
    
    # Assign position to each item in original list
    # Return as float
    ypos = [ydict[x] for x in ylabels]
    ypos = np.float64(ypos)
    
    return (ylabels, ypos, ydict)
    
def getColourBarMinMax(minSamples, maxSamples):
    """ Determine top of Color Bar """
    
    # Get whether 10s, 100s, 1000s of samples
    powerMax = math.floor(math.log10(maxSamples))
    powerMin = math.floor(math.log10(minSamples))
    
    # If 100s or 1000s or samples, round to nearest 100
    # If 10s of samples, round to nearest 10
    if powerMax > 1:
        roundMax = 10**2
        maxBar = math.ceil(maxSamples/roundMax) * roundMax
    elif powerMin == 1:
        maxBar = math.ceil(maxSamples/10) * 10    
    
    if powerMin > 1:
        roundMin = 10**2
        minBar = math.floor(minSamples/roundMin) * roundMin
    elif powerMin == 1:
        minBar = math.floor(minSamples/10) * 10

    return (minBar, maxBar)

def getPlotColours(colMap,sampleNumbers,bottom,top):
    """ Return a list of colours based on relative
        position in colour map """
    
    # Colour map is row 0 to end (light to dark)
    # Get difference between colour bar limits
    sampleDiff = top-bottom
    
    # Get relative positions of rows based on number of samples
    #(difference between sample number and lower limit) as proportion of
    # rows in colour map array
    sampleRows = ((sampleNumbers - bottom) / sampleDiff) * len(colMap)
    
    # round sample row down, take minimum of this and length of colour map array
    # Ensures a number between 0 and length of array
    sampleRows = [min(math.floor(x),len(colMap)-1) for x in sampleRows]
    
    # return RGB from row
    colours = [(colMap[x,0],colMap[x,1],colMap[x,2]) for x in sampleRows]
    return colours


def getPlotBars(y,start,end,colours,progress):
    """ Get all axes properties and return axes """
    
    # Create figure and axes (colour bar 30th of width)
    fig, (ax,ax2) = plt.subplots(1,2,figsize=(12, 7.5), gridspec_kw = {'width_ratios':[3, 0.1]},sharey=False)

    # Create main bars for study start and end
    ax.barh(y, end - start, left=start, height=0.3, align='center',
            alpha=1,color=colours)
    
    # Create progress bar
    ax.barh(y+0.15, (end - start)*progress, left=start, height=-0.08, align='edge',
            alpha=1,color='darkorange',label='% Complete')

    return fig, ax, ax2

def getPlotProperties(ax,yLabelsPos,minDate=None,maxDate=None):
    """ Set axis properties for Gantt chart """
    
    # Tight axis and legend
    ax.axis('tight')
    ax.legend()

    # Add grid in x-direction only
    ax.grid(color = 'darkgray', linestyle = ':')
    ax.yaxis.grid(False)

    # x axis as date
    ax.xaxis_date()
    ax.margins(x=0)
    
    
    # x-axis format
    myFmt = mpdt.DateFormatter("%b-%y")
    ax.xaxis.set_major_formatter(myFmt)
    ax.xaxis.set_major_locator(mpdt.MonthLocator(interval=2))
    

    # x-axis tick marks    
    labelsx = ax.get_xticklabels()
    plt.setp(labelsx, rotation=30, fontsize=10)
    
    # x-axis limits
    if minDate is None or maxDate is None:    
        
        # Extend lower to start of month and upper to end of month
        # Default x axis
        xmin, xmax = setAutoDateRange(ax.get_xlim())        
        ax.set_xlim(xmin=xmin,xmax=xmax)
        
    elif minDate is not None and maxDate is not None:
        
        # Set the date range from user input
        # Get dates as floats
        xlims = [mpdt.datestr2num(minDate),mpdt.datestr2num(maxDate)]
        
        # Tidy up axis limits
        
        xmin, xmax = setAutoDateRange(xlims)
        ax.set_xlim(xmin=xmin,xmax=xmax)
    
    ## Set y ticks
    # Get unique values for y position and labels (from dictionary)
    ylabels = list(yLabelsPos.keys())
    ypos = list(yLabelsPos.values())
    ax.set_yticks(ypos)
    ax.set_yticklabels(ylabels)
    ax.tick_params(axis='y', labelsize=10)

    # Set y axis limits
    ymax = max(ypos)+0.5
    ax.set_ylim(ymin = -0.1, ymax = ymax)
    
    return ax


def getPlotColourBar(ax,colMap,minTick,maxTick):
    """ Create a custom colour bar to place in axis 2 """
    
    # Create custom colour map
    cm = mpl.colors.ListedColormap(colMap)
    
    # Normalise to number of samples
    norm = mpl.colors.Normalize(vmin=minTick, vmax=maxTick)
    
    # Create colourbar
    cbar = mpl.colorbar.ColorbarBase(ax=ax, cmap=cm,norm=norm)
    
    ## Tick marks
    # 6 evenly spaced ticks
    tickSpace = (maxTick-minTick)/5
    
    # Because 6 labels but 5 spaces, add tickspace at the top
    tcks=arange(minTick,maxTick+tickSpace,tickSpace)
    cbar.set_ticks(tcks)
    cbar.set_label('No. Samples',fontsize=12,labelpad=-80)
    cbar.ax.set_yticklabels([str(a) for a in tcks],fontsize=11)

    return(cbar)
#    
def setAutoDateRange(xlimits):
    """ Pass ax_xlim and 
        set range to start of first month and end of last month.  """
    
    # Get smallest start date and largest end date as float
    minDate = xlimits[0]
    maxDate = xlimits[1]
    
    # Return first of month
    minYear = int(mpdt.num2date(minDate).year)
    minMonth = int(mpdt.num2date(minDate).month)
    
    # x minimum
    xminDate = dt.datetime(minYear,minMonth,int(1))
    xminDate = mpdt.date2num(xminDate)
    
    # Get last day of month
    maxYear = int(mpdt.num2date(maxDate).year)
    maxMonth = int(mpdt.num2date(maxDate).month)
    lastDay = calendar.monthrange(maxYear,maxMonth)[1]
    
    # x max (add 1 to get last tick mark)
    xmaxDate = dt.datetime(maxYear,maxMonth,lastDay)
    xmaxDate = mpdt.date2num(xmaxDate)+1
    
    return (xminDate,xmaxDate)
    
    