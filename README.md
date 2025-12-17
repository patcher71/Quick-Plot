# Quick-Plot

This program was designed to quickly review electrophysiological recording data performed with the WinLTP software (https://winltp.co.uk/) developed by Dr. Bill Anderson. 
After selecting the directory containing the data, the program features three tabs summarized as follows:

## Sweep Select

This tab allows you to view individual sweep traces (currently, .P0 or .P1 files) and quickly create sweep averages. Up to 4 different averages may be created containing any number of sweeps, and these can be named according to user-defined periods (for example, control, drug, etc.) The average files can then be saved to an Excel spreadsheet for plotting in other programs.

<img width="1548" height="1157" alt="image" src="https://github.com/user-attachments/assets/682ab009-3585-4e8f-89dd-4252a3d3f1b1" />


### Time Plot

This tab reads in the Microsoft Excel (XLS) 'Ampfile' workbook created during either online recording/analysis or following reanalysis of sweeps using the WinLTP reanalysis package. The selected workbook file will be loaded into the main window, and the time base for the x-axis can be selected (index, Time_min, or Time_s). The user can also choose to filter by stimulation (S0 or S1) in the event that the user wants to plot separate figures based on the type of stimulation (either two stimulators, or perhaps electrical vs. optical, etc).

The user also can select to 'Filter Sweeps to current workbook' which will then limit all sweeps visible on the 'Sweeps Select' and 'Drug Markers' tab to those in the current workbook. 

The user can then choose 4 different plots. Typically, for intracellular recordings, this might be peak amplitude, DC (holding current), Rs, and Rm.  But in essence, any of the parameters found in the last columns of the workbook can be plotted. In the event that paired pulse data are obtained, a 5th plot will calculate the paired pulse ratio (peak amp 2/peak amp 1).  All plots can be saved to an Excel workbook for further graphing/analysis, or as PNG image files.

<img width="1542" height="1137" alt="image" src="https://github.com/user-attachments/assets/165528d6-caa9-4bbf-93d1-b56a8df63e92" />


#### Drug Markers

This feature was added so that the user can quickly add indicators to the plots (and to the workbook) indicating which sweeps drugs were applied.  Up to 4 drug on/drug off options can be selected. Simply select one sweep in each box corresponding to when drugs were turned on or off. Markers will be added to the plot, as well as to any saved workbooks, for further analysis.

<img width="1575" height="1152" alt="image" src="https://github.com/user-attachments/assets/81b5ecad-7181-4eb0-b494-16787a52fd38" />
<img width="1523" height="1153" alt="image" src="https://github.com/user-attachments/assets/5ca05df0-4636-4a0e-9f3b-ed879708d619" />
