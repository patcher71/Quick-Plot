classdef QuickPlots_v101 < matlab.apps.AppBase

    % Properties that correspond to app components
    properties (Access = public)
        QuickPlotterUIFigure         matlab.ui.Figure
        DirectoryEditField           matlab.ui.control.EditField
        DirectoryEditFieldLabel      matlab.ui.control.Label
        ChooseDirectoryButton        matlab.ui.control.Button
        TabGroup                     matlab.ui.container.TabGroup
        SweepSelectTab               matlab.ui.container.Tab
        SweepExtDropDown             matlab.ui.control.DropDown
        SweepFileExtensionTypeDropDownLabel  matlab.ui.control.Label
        ClearAVG4Button              matlab.ui.control.Button
        ClearAVG3Button              matlab.ui.control.Button
        ClearAVG2Button              matlab.ui.control.Button
        ClearAVG1Button              matlab.ui.control.Button
        AVG4NameEditField            matlab.ui.control.EditField
        AVG4NameEditFieldLabel       matlab.ui.control.Label
        AVG3NameEditField            matlab.ui.control.EditField
        AVG3NameEditFieldLabel       matlab.ui.control.Label
        AVG2NameEditField            matlab.ui.control.EditField
        AVG2NameEditFieldLabel       matlab.ui.control.Label
        AVG1NameEditField            matlab.ui.control.EditField
        AVG1NameEditFieldLabel       matlab.ui.control.Label
        AddtoAVG4Button              matlab.ui.control.Button
        AddtoAVG3Button              matlab.ui.control.Button
        AddtoAVG2Button              matlab.ui.control.Button
        AddtoAVG1Button              matlab.ui.control.Button
        FileListBox                  matlab.ui.control.ListBox
        FileListBoxLabel             matlab.ui.control.Label
        SaveAVGtoTXTfileButton       matlab.ui.control.Button
        SweepFileAxes                matlab.ui.control.UIAxes
        AVGFileAxes                  matlab.ui.control.UIAxes
        TimePlotTab                  matlab.ui.container.Tab
        StimFilterDropDown           matlab.ui.control.DropDown
        StimFilterDropDownLabel      matlab.ui.control.Label
        ExportTimePlotsButton        matlab.ui.control.Button
        PkAmpParamDropDown           matlab.ui.control.DropDown
        Graph4Label                  matlab.ui.control.Label
        RmParamDropDown              matlab.ui.control.DropDown
        Graph3Label                  matlab.ui.control.Label
        RsParamDropDown              matlab.ui.control.DropDown
        Graph2Label                  matlab.ui.control.Label
        DCParamDropDown              matlab.ui.control.DropDown
        Graph1Label                  matlab.ui.control.Label
        FilterP0ByWorkbookCheckBox   matlab.ui.control.CheckBox
        SavePlotstoWorkbookButton    matlab.ui.control.Button
        TimeBaseDropDown             matlab.ui.control.DropDown
        TimeBaseforXaxisLabel        matlab.ui.control.Label
        WorkbookUITable              matlab.ui.control.Table
        WorkbookListBox              matlab.ui.control.ListBox
        WorkbookListBoxLabel         matlab.ui.control.Label
        RatioAxes                    matlab.ui.control.UIAxes
        RmAxes                       matlab.ui.control.UIAxes
        RsAxes                       matlab.ui.control.UIAxes
        PkAmpAxes                    matlab.ui.control.UIAxes
        DCAxes                       matlab.ui.control.UIAxes
        DrugMarkersTab               matlab.ui.container.Tab
        Drug4EditField               matlab.ui.control.EditField
        DrugEditField_2Label_3       matlab.ui.control.Label
        Drug3EditField               matlab.ui.control.EditField
        DrugEditField_2Label_2       matlab.ui.control.Label
        Drug2EditField               matlab.ui.control.EditField
        Drug2Label                   matlab.ui.control.Label
        Drug1EditField               matlab.ui.control.EditField
        Drug1EditFieldLabel          matlab.ui.control.Label
        Drug4OffListBox              matlab.ui.control.ListBox
        DrugOffListBox_4Label        matlab.ui.control.Label
        Drug4OnListBox               matlab.ui.control.ListBox
        DrugOnListBox_4Label         matlab.ui.control.Label
        Drug3OffListBox              matlab.ui.control.ListBox
        DrugOffListBox_3Label        matlab.ui.control.Label
        Drug3OnListBox               matlab.ui.control.ListBox
        DrugOnListBox_3Label         matlab.ui.control.Label
        Drug2OffListBox              matlab.ui.control.ListBox
        DrugOffListBox_2Label        matlab.ui.control.Label
        Drug2OnListBox               matlab.ui.control.ListBox
        DrugOnListBox_2Label         matlab.ui.control.Label
        Drug1OffListBox              matlab.ui.control.ListBox
        DrugOffListBoxLabel          matlab.ui.control.Label
        Drug1OnListBox               matlab.ui.control.ListBox
        DrugOnListBoxLabel           matlab.ui.control.Label
        SweepReanalysisTab           matlab.ui.container.Tab
        ToggleIncludeButton          matlab.ui.control.Button
        ApplyReanalysisPeaksButton   matlab.ui.control.Button
        ApplyReanalysisRmButton      matlab.ui.control.Button
        Peak2EndEditField            matlab.ui.control.NumericEditField
        Peak2EndmsEditFieldLabel     matlab.ui.control.Label
        Peak2StartEditField          matlab.ui.control.NumericEditField
        Peak2StartmsEditFieldLabel   matlab.ui.control.Label
        SaveTabletoExcelButton       matlab.ui.control.Button
        AddAlltoTableButton          matlab.ui.control.Button
        RmWindowEndEditField         matlab.ui.control.NumericEditField
        RmMeasureStartmsLabel_2      matlab.ui.control.Label
        RmWindowStartEditField       matlab.ui.control.NumericEditField
        RmMeasureStartmsLabel        matlab.ui.control.Label
        RmBaselineEndEditField       matlab.ui.control.NumericEditField
        RmBaselineEndmsLabel         matlab.ui.control.Label
        RmBaselineStartEditField     matlab.ui.control.NumericEditField
        RmBaselineStartmsEditFieldLabel  matlab.ui.control.Label
        ReanalysisTable              matlab.ui.control.Table
        PeakAvgPointsEditField       matlab.ui.control.NumericEditField
        ptsforPeakavgEditFieldLabel  matlab.ui.control.Label
        RmResultLabel                matlab.ui.control.Label
        PeakResultLabel              matlab.ui.control.Label
        RecomputeButton              matlab.ui.control.Button
        Peak1EndEditField            matlab.ui.control.NumericEditField
        WindowEndmsLabel             matlab.ui.control.Label
        Peak1StartEditField          matlab.ui.control.NumericEditField
        WindowstartmsLabel           matlab.ui.control.Label
        BaselineEndEditField         matlab.ui.control.NumericEditField
        BaselineEndmsLabel           matlab.ui.control.Label
        BaselineStartEditField       matlab.ui.control.NumericEditField
        BaselineStartmsLabel         matlab.ui.control.Label
        SelectSweepListBox           matlab.ui.control.ListBox
        SelectSweepListBoxLabel      matlab.ui.control.Label
        UIAxes                       matlab.ui.control.UIAxes
    end

    
    properties (Access = private)
    P0Dir string               % current directory
    P0Files struct            % output of dir()
    WorkbookFiles struct
    CurrentWorkbook table
    FilterP0ByWorkbook logical = false;

    LastSweepX double = []    % last single sweep X
    LastSweepY double = []    % last single sweep Y
    LastSweepUnit string = "pA";   % 'pA' or 'mV'
    TimePlotUnit string = ""   % e.g. "1 pA" or "1 mV" from workbook
    TimePlotParamVars cell = {}   % e.g. {'DC','PkAmp','Rs_Mohm','Rm_Mohm'}
    
    % NEW: four separate averages

    AvgX1 double = [];   AvgY1 double = [];   AvgCount1 double = 0;
    AvgX2 double = [];   AvgY2 double = [];   AvgCount2 double = 0;
    AvgX3 double = [];   AvgY3 double = [];   AvgCount3 double = 0;
    AvgX4 double = [];   AvgY4 double = [];   AvgCount4 double = 0;

    %existing properties
    DrugMarkers struct   % 1x4 struct array with name/onFile/offFile

    ReanX double = []          % time (ms) of current reanalysis sweep
    ReanY double = []          % AD0 trace of current reanalysis sweep
    ReanUnit string = ""       % 'pA' or 'mV'

    BaselineStart double = 0
    BaselineEnd   double = 5
    Peak1Start  double = 10
    Peak1End    double = 20
    Peak2Start double = NaN
    Peak2End double = NaN

    % NEW: Rm-specific windows
    RmBaselineStart double = 0
    RmBaselineEnd   double = 0
    RmWindowStart   double = 0
    RmWindowEnd     double = 0

    BaselineLines  = gobjects(1,2)
    Peak1Lines  = gobjects(1,2)
    Peak2Lines  = gobjects(1,2)

    % NEW: Rm cursor lines
    RmBaselineLines = gobjects(1,2)
    RmWindowLines   = gobjects(1,2)


    ReanMeta struct = struct()   % metadata for current reanalysis sweep
    
    % store latest metrics for table
    LastReanPeak1 double = NaN
    LastReanPeak1Time double = NaN
    LastReanPeak2 double = NaN
    LastReanPeak2Time double = NaN
    LastReanRm double = NaN

    ReanWindowsInitialized logical = false;


    
    end
    
    methods (Access = private)
        
        function [t_ms, y_AD0, unitStr,Vm_mV, meta] = readP0File(app, filename)
  
            %READP0FILE  Read WinLTP .P0 file.
%   Returns:
%     t_ms   : time vector (ms)
%     y_AD0  : AD0 data
%     unitStr: AD0 unit ('pA' or 'mV')
%     Vm_mV  : holding potential from AD1(1) (mV)
%     meta   : struct with NumSamples & SampleInterval_ms

    fid = fopen(filename, 'r');
    if fid == -1
        error('Could not open file: %s', filename);
    end
    cleaner = onCleanup(@() fclose(fid));

    NumSamples        = NaN;
    SampleInterval_ms = NaN;
    unitStr           = "pA";   % default
    RsRmPulseAmp_mV = NaN;   % <-- NEW
    dataSectionReached = false;

    % ---------- Parse header ----------
    while true
        tline = fgetl(fid);
        if ~ischar(tline)
            break;  % EOF
        end

        % Number of samples
        if contains(tline, '"NumSamples"')
            tok = regexp(tline, '"NumSamples"\s+(\d+)', 'tokens', 'once');
            if ~isempty(tok)
                NumSamples = str2double(tok{1});
            end

        % Sample interval
        elseif contains(tline, '"SampleInterval_ms"')
            tok = regexp(tline, '"SampleInterval_ms"\s+([-\d\.Ee+]+)', ...
                         'tokens', 'once');
            if ~isempty(tok)
                SampleInterval_ms = str2double(tok{1});
            end

           elseif contains(tline, '"RsRmPulseAmp_mV_mV"')
                % e.g. "RsRmPulseAmp_mV_mV"   -10.0   0.0
                tok = regexp(tline, '"RsRmPulseAmp_mV_mV"\s+([-\d\.Ee+]+)', ...
                             'tokens', 'once');
                if ~isempty(tok)
                    RsRmPulseAmp_mV = str2double(tok{1});   % first (IC0) value
                end 

        % AD0 data header
        elseif startsWith(strtrim(tline), '"AD0"')
            % Next line is units, e.g.  "pA"   "mV"
            unitLine = fgetl(fid);
            if ischar(unitLine)
                tokens = regexp(unitLine, '"([^"]+)"', 'tokens');
                if ~isempty(tokens)
                    unitStr = string(tokens{1}{1});   % 'pA' or 'mV'
                end
            end
            dataSectionReached = true;
            break;
        end
    end

    if ~dataSectionReached
        error('Could not find AD0 data block in %s', filename);
    end

    if isnan(SampleInterval_ms)
        error('SampleInterval_ms not found in header for %s', filename);
    end

    % ---------- Read numeric data (AD0, AD1) ----------
    data = textscan(fid, '%f%f%*[^\n]', 'CollectOutput', true);
    data = data{1};
    y_AD0  = data(:,1);   % AD0
    AD1_mV = data(:,2);   % AD1

    if ~isnan(NumSamples)
        y_AD0  = y_AD0(1:NumSamples);
        AD1_mV = AD1_mV(1:NumSamples);
    else
        NumSamples = numel(y_AD0);
    end

    t_ms = (0:NumSamples-1).' * SampleInterval_ms;

    % Holding potential from first AD1 sample
    Vm_mV = AD1_mV(1);

    % Metadata struct
    meta = struct();
    meta.NumSamples        = NumSamples;
    meta.SampleInterval_ms = SampleInterval_ms;
    meta.Vm_mV             = Vm_mV;
    meta.RsRmPulseAmp_mV   = RsRmPulseAmp_mV;   % <-- NEW
   end
   

        function writeAVGTextFile(app, filename, t_ms, y_pA)
    % Writes a simple 2-column text file:
    % time_ms    avg_AD0_pA

    fid = fopen(filename, 'w');
    if fid == -1
        error('Could not create file: %s', filename);
    end

    cleaner = onCleanup(@() fclose(fid));

    % Header line
    fprintf(fid, 'Time_ms\tAD0_pA\n');

    % Body
    for i = 1:numel(t_ms)
        fprintf(fid, '%f\t%f\n', t_ms(i), y_pA(i));
    end
        end

function addCurrentSweepToAverage(app, idx)
     % Need a current sweep loaded
    if isempty(app.LastSweepX)
        return;
    end

    x = app.LastSweepX;
    y = app.LastSweepY;

    % ----- BASELINE NORMALIZATION: first 10 ms -----
    % logical index of samples up to 10 ms
    baselineIdx = x <= 10;          % hard-coded 10 ms window

    if any(baselineIdx)
        baseline = mean(y(baselineIdx));
    else
        baseline = 0;               % safety fallback
    end

    y = y - baseline;               % subtract baseline so trace starts ~0
    % ------------------------------------------------

    % Select the right average group
    switch idx
        case 1
            X = app.AvgX1; Y = app.AvgY1; n = app.AvgCount1;
        case 2
            X = app.AvgX2; Y = app.AvgY2; n = app.AvgCount2;
        case 3
            X = app.AvgX3; Y = app.AvgY3; n = app.AvgCount3;
        case 4
            X = app.AvgX4; Y = app.AvgY4; n = app.AvgCount4;
        otherwise
            return;
    end

    if n == 0
        % First sweep for this group
        X = x;
        Y = y;
        n = 1;
    else
        % Check that time vectors match
        if numel(x) ~= numel(X) || any(x ~= X)
            % Could interpolate here; for now ignore mismatched sweep
            return;
        end

        % Running average
        Y = (Y * n + y) / (n + 1);
        n = n + 1;
    end

    % Write back to the appropriate properties
    switch idx
        case 1
            app.AvgX1 = X; app.AvgY1 = Y; app.AvgCount1 = n;
        case 2
            app.AvgX2 = X; app.AvgY2 = Y; app.AvgCount2 = n;
        case 3
            app.AvgX3 = X; app.AvgY3 = Y; app.AvgCount3 = n;
        case 4
            app.AvgX4 = X; app.AvgY4 = Y; app.AvgCount4 = n;
    end

    % Update overlay plot
    app.updateAverageOverlayPlot();
end

function [legendLabel, headerBase, count] = getAvgName(app, idx)
    % Returns:
    %  legendLabel:  text for plot legend, e.g. 'Baseline (3 sweeps)'
    %  headerBase :  base for file header, e.g. 'Baseline'
    %  count      :  sweep count for this average

    switch idx
        case 1
            rawName    = strtrim(app.AVG1NameEditField.Value);
            defaultName = 'AVG1';
            count      = app.AvgCount1;
        case 2
            rawName    = strtrim(app.AVG2NameEditField.Value);
            defaultName = 'AVG2';
            count      = app.AvgCount2;
        case 3
            rawName    = strtrim(app.AVG3NameEditField.Value);
            defaultName = 'AVG3';
            count      = app.AvgCount3;
        case 4
            rawName    = strtrim(app.AVG4NameEditField.Value);
            defaultName = 'AVG4';
            count      = app.AvgCount4;
        otherwise
            rawName    = '';
            defaultName = 'AVG';
            count      = 0;
    end

    if isempty(rawName)
        base = defaultName;
    else
        base = rawName;
    end

    legendLabel = sprintf('%s (%d sweeps)', base, count);
    headerBase  = regexprep(base, '\s+', '_');   % replace spaces with underscores
end


function writeAllAveragesTextFile(app, filename)

    groups = struct('legendLabel', {}, 'headerBase', {}, ...
                    'count', {}, 'X', {}, 'Y', {});

    if app.AvgCount1 > 0
        [legendLabel, headerBase, count] = app.getAvgName(1);
        groups(end+1) = struct('legendLabel', legendLabel, ...
                               'headerBase', headerBase, ...
                               'count', count, ...
                               'X', app.AvgX1, ...
                               'Y', app.AvgY1);
    end
    if app.AvgCount2 > 0
        [legendLabel, headerBase, count] = app.getAvgName(2);
        groups(end+1) = struct('legendLabel', legendLabel, ...
                               'headerBase', headerBase, ...
                               'count', count, ...
                               'X', app.AvgX2, ...
                               'Y', app.AvgY2);
    end
    if app.AvgCount3 > 0
        [legendLabel, headerBase, count] = app.getAvgName(3);
        groups(end+1) = struct('legendLabel', legendLabel, ...
                               'headerBase', headerBase, ...
                               'count', count, ...
                               'X', app.AvgX3, ...
                               'Y', app.AvgY3);
    end
    if app.AvgCount4 > 0
        [legendLabel, headerBase, count] = app.getAvgName(4);
        groups(end+1) = struct('legendLabel', legendLabel, ...
                               'headerBase', headerBase, ...
                               'count', count, ...
                               'X', app.AvgX4, ...
                               'Y', app.AvgY4);
    end

    if isempty(groups)
        error('writeAllAveragesTextFile:NoData', 'No averages to write.');
    end

    % Use time base from first group
    Time = groups(1).X;
    n    = numel(Time);
    for k = 2:numel(groups)
        if numel(groups(k).X) ~= n || any(groups(k).X ~= Time)
            error('Time vectors do not match between average groups.');
        end
    end

    % Use the current sweep unit for column labels
    unitStr = char(app.LastSweepUnit);
    if isempty(unitStr)
        unitStr = 'pA';  % fallback
    end

    fid = fopen(filename, 'w');
    if fid == -1
        error('Could not create file: %s', filename);
    end
    cleaner = onCleanup(@() fclose(fid));

    % ---- Header line ----
    fprintf(fid, 'Time_ms');
    for k = 1:numel(groups)
        fprintf(fid, '\t%s_%dsweeps_%s', ...
            groups(k).headerBase, ...
            groups(k).count, ...
            unitStr);                    % <-- was hard-coded 'pA'
    end
    fprintf(fid, '\n');

    % ---- Data lines ----
    for i = 1:n
        fprintf(fid, '%f', Time(i));
        for k = 1:numel(groups)
            fprintf(fid, '\t%f', groups(k).Y(i));
        end
        fprintf(fid, '\n');
    end
end




function updateAverageOverlayPlot(app)

   ax = app.AVGFileAxes;
    cla(ax);
    hold(ax, 'on');

    if app.AvgCount1 > 0
        [lbl, ~, ~] = app.getAvgName(1);
        plot(ax, app.AvgX1, app.AvgY1, 'DisplayName', lbl);
    end
    if app.AvgCount2 > 0
        [lbl, ~, ~] = app.getAvgName(2);
        plot(ax, app.AvgX2, app.AvgY2, 'DisplayName', lbl);
    end
    if app.AvgCount3 > 0
        [lbl, ~, ~] = app.getAvgName(3);
        plot(ax, app.AvgX3, app.AvgY3, 'DisplayName', lbl);
    end
    if app.AvgCount4 > 0
        [lbl, ~, ~] = app.getAvgName(4);
        plot(ax, app.AvgX4, app.AvgY4, 'DisplayName', lbl);
    end

    hold(ax, 'off');

    xlabel(ax, 'Time (ms)');
    ylabel(ax, sprintf('AD0 (%s)', app.LastSweepUnit));
    title(ax, 'Overlay of averages');

    if app.AvgCount1 + app.AvgCount2 + app.AvgCount3 + app.AvgCount4 > 0
        legend(ax, 'Location', 'best');
    end
end

function clearAverage(app, idx)
    switch idx
        case 1
            app.AvgX1 = [];
            app.AvgY1 = [];
            app.AvgCount1 = 0;
            %app.AVG1NameEditField.Value = '';   % optional, remove if you want to keep the name
        case 2
            app.AvgX2 = [];
            app.AvgY2 = [];
            app.AvgCount2 = 0;
            %app.AVG2NameEditField.Value = '';
        case 3
            app.AvgX3 = [];
            app.AvgY3 = [];
            app.AvgCount3 = 0;
            %app.AVG3NameEditField.Value = '';
        case 4
            app.AvgX4 = [];
            app.AvgY4 = [];
            app.AvgCount4 = 0;
            %app.AVG4NameEditField.Value = '';
    end

    % Redraw overlay without this average
    app.updateAverageOverlayPlot();
end

function updateTimePlots(app)

          %----------- Start from full workbook -----------
    Tfull = app.CurrentWorkbook;

    if isempty(Tfull) || height(Tfull) == 0
        cla(app.DCAxes); cla(app.RsAxes);
        cla(app.PkAmpAxes); cla(app.RmAxes);
        if isprop(app,'RatioAxes'); cla(app.RatioAxes); end
        return;
    end

    varsFull = Tfull.Properties.VariableNames;

    %----------- Apply Include flag (if present) -----------
    if ismember("Include", varsFull)
        maskInclude = Tfull.Include ~= 0 & ~isnan(Tfull.Include);
    else
        maskInclude = true(height(Tfull),1);
    end

    T = Tfull(maskInclude,:);
    if height(T) == 0
        cla(app.DCAxes); cla(app.RsAxes);
        cla(app.PkAmpAxes); cla(app.RmAxes);
        if isprop(app,'RatioAxes'); cla(app.RatioAxes); end
        return;
    end

    %----------- Now work with filtered T -----------
    vars = T.Properties.VariableNames;

    % --------- Filter rows by filename extension (P0/P1) ----------
    mask = true(height(T),1);

    if ismember('Filename_Ext', vars) && isprop(app,'SweepExtDropDown')
        switch app.SweepExtDropDown.Value
            case 'P0 only'
                mask = mask & endsWith(string(T.Filename_Ext), '.P0', 'IgnoreCase', true);
            case 'P1 only'
                mask = mask & endsWith(string(T.Filename_Ext), '.P1', 'IgnoreCase', true);
            otherwise
                % 'All' -> no extra restriction
        end
    end

    % --------- Filter by Sx (stimulus 0 / 1) ----------
    if ismember('Sx', vars) && isprop(app,'StimFilterDropDown')
        switch app.StimFilterDropDown.Value
            case 'S0'
                mask = mask & (T.Sx == 0);
            case 'S1'
                mask = mask & (T.Sx == 1);
            otherwise
                % 'All' -> no further restriction
        end
    end

    % Apply combined mask
    T = T(mask,:);
    if height(T) == 0
        cla(app.DCAxes); cla(app.RsAxes);
        cla(app.PkAmpAxes); cla(app.RmAxes);
        if isprop(app,'RatioAxes'); cla(app.RatioAxes); end
        return;
    end

    % refresh vars after subsetting
    vars = T.Properties.VariableNames;

    % --------- X axis selection ----------
    switch app.TimeBaseDropDown.Value
        case 'Time_min'
            if ismember('Time_min', vars)
                x = T.Time_min;
                xLabel = 'Time (min)';
            else
                x = (1:height(T))';
                xLabel = 'Sweep #';
            end

        case 'Time_sec'
            if ismember('Time_sec', vars)
                x = T.Time_sec;
                xLabel = 'Time (s)';
            else
                x = (1:height(T))';
                xLabel = 'Sweep #';
            end

        otherwise  % 'Index'
            x = (1:height(T))';
            xLabel = 'Sweep #';
    end

    % --------- Pulse number vector ----------
    pulVar = vars(startsWith(vars,'Pul'));
    if ~isempty(pulVar)
        pul = T.(pulVar{1});
    else
        pul = ones(height(T),1);  % fallback: treat all as pulse 1
    end

    % --------- Use dropdowns to choose parameters ----------
    dcParam    = string(app.DCParamDropDown.Value);
    rsParam    = string(app.RsParamDropDown.Value);
    rmParam    = string(app.RmParamDropDown.Value);
    pkampParam = string(app.PkAmpParamDropDown.Value);

    % DC
    app.plotTimeParamOnAxis(app.DCAxes,   dcParam,    x, xLabel, pul, T);

    % Rs
    app.plotTimeParamOnAxis(app.RsAxes,   rsParam,    x, xLabel, pul, T);

    % Rm
    app.plotTimeParamOnAxis(app.RmAxes,   rmParam,    x, xLabel, pul, T);

    % Peak amplitude
    app.plotTimeParamOnAxis(app.PkAmpAxes, pkampParam, x, xLabel, pul, T);

    % --------- Optional P2/P1 ratio ----------
    if isprop(app,'RatioAxes') && ismember('PkAmp',vars) && any(pul == 2)

        if ismember('Filename', vars)
            filename = T.Filename;
        else
            filename = (1:height(T))';   % fallback
        end

        [G,~,gid] = unique(filename);
        nSweeps   = numel(G);
        ratioX    = nan(nSweeps,1);
        ratioY    = nan(nSweeps,1);

        for k = 1:nSweeps
            rows = gid == k;
            idx1 = rows & pul == 1;
            idx2 = rows & pul == 2;

            if any(idx1) && any(idx2)
                p1 = T.PkAmp(find(idx1,1,'first'));
                p2 = T.PkAmp(find(idx2,1,'first'));

                if p1 ~= 0 && ~isnan(p1) && ~isnan(p2)
                    ratioY(k) = p2 / p1;
                    ratioX(k) = mean(x(rows));   % average X for that sweep
                end
            end
        end

        valid = ~isnan(ratioY);
        cla(app.RatioAxes);
        plot(app.RatioAxes, ratioX(valid), ratioY(valid));
        xlabel(app.RatioAxes, xLabel);
        ylabel(app.RatioAxes, 'P2 / P1');
        title(app.RatioAxes, 'Pulse2 / Pulse1 ratio');
        app.drawDrugMarkersOnAxes(app.RatioAxes, x, T);
    end
end

function syncDrugMarkersFromUI(app)
    names = {app.Drug1EditField.Value, app.Drug2EditField.Value, ...
             app.Drug3EditField.Value, app.Drug4EditField.Value};

    onVals  = {app.Drug1OnListBox.Value,  app.Drug2OnListBox.Value, ...
               app.Drug3OnListBox.Value,  app.Drug4OnListBox.Value};

    offVals = {app.Drug1OffListBox.Value, app.Drug2OffListBox.Value, ...
               app.Drug3OffListBox.Value, app.Drug4OffListBox.Value};

    n = 4;
    markers = struct('name', strings(1,n), ...
                     'onFile', strings(1,n), ...
                     'offFile', strings(1,n));

    for k = 1:n
        nm = strtrim(string(names{k}));
        if nm == ""
            nm = "Drug " + k;
        end
        markers(k).name   = nm;
        markers(k).onFile  = string(onVals{k});
        markers(k).offFile = string(offVals{k});
    end

    app.DrugMarkers = markers;
end

function drawDrugMarkersOnAxes(app, ax, x, T)
    % x: x-axis vector used for this plot
    % T: current workbook table

    if isempty(app.DrugMarkers)
        return;
    end

    vars = T.Properties.VariableNames;

    % Use Filename_Ext if present, else Filename
    fnVar = '';
    if ismember('Filename_Ext', vars)
        fnVar = 'Filename_Ext';
    elseif ismember('Filename', vars)
        fnVar = 'Filename';
    else
        return;
    end

    fnCol = string(T.(fnVar));

    colors = lines(4);   % 4 distinct colors

    hold(ax, 'on');
    for k = 1:numel(app.DrugMarkers)
        dm = app.DrugMarkers(k);

        if dm.onFile ~= ""
            idxOn = find(fnCol == dm.onFile, 1);
            if ~isempty(idxOn)
                xOn = x(idxOn);
                xline(ax, xOn, '--', 'Color', colors(k,:), ...
                      'Label', sprintf('%s on', dm.name), ...
                      'LabelOrientation', 'horizontal');
            end
        end

        if dm.offFile ~= ""
            idxOff = find(fnCol == dm.offFile, 1);
            if ~isempty(idxOff)
                xOff = x(idxOff);
                xline(ax, xOff, ':', 'Color', colors(k,:), ...
                      'Label', sprintf('%s off', dm.name), ...
                      'LabelOrientation', 'horizontal');
            end
        end
    end
    hold(ax, 'off');
end


function [summaryTbl, ratioTbl, drugTbl, xLabel] = buildTimePlotExportTables(app)

    T = app.CurrentWorkbook;
    summaryTbl = table();
    ratioTbl   = table();
    drugTbl    = table();
    xLabel     = '';

    if isempty(T) || height(T) == 0
        return;
    end

    vars = T.Properties.VariableNames;

    % ---------- X axis ----------
    switch app.TimeBaseDropDown.Value
        case 'Time_min'
            if ismember('Time_min', vars)
                x = T.Time_min;
                xLabel = 'Time_min';
            else
                x = (1:height(T))';
                xLabel = 'Index';
            end
        case 'Time_sec'
            if ismember('Time_sec', vars)
                x = T.Time_sec;
                xLabel = 'Time_sec';
            else
                x = (1:height(T))';
                xLabel = 'Index';
            end
        otherwise
            x = (1:height(T))';
            xLabel = 'Index';
    end

    % ---------- Pulse column ----------
    pulVar = vars(startsWith(vars, 'Pul'));   % e.g. 'Pul_'
    if ~isempty(pulVar)
        pul = T.(pulVar{1});
    else
        pul = ones(height(T),1);
    end

    n = height(T);

    % ---------- Full-length series (NaN where not defined) ----------
    DC  = nan(n,1);
    Rs  = nan(n,1);
    Rm  = nan(n,1);
    P1  = nan(n,1);    % PkAmp pulse 1
    P2  = nan(n,1);    % PkAmp pulse 2

    if ismember('DC', vars)
         idx1 = pul == 1;
         DC(idx1) = T.DC(idx1);   % only pulse 1; others stay NaN
    end
    if ismember('Rs_Mohm', vars)
        idx1 = pul == 1;
        Rs(idx1) = T.Rs_Mohm(idx1);
    end
    if ismember('Rm_Mohm', vars)
        idx1 = pul == 1;
        Rm(idx1) = T.Rm_Mohm(idx1);
    end
    if ismember('PkAmp', vars)
        idx1 = pul == 1;
        idx2 = pul == 2;
        P1(idx1) = T.PkAmp(idx1);
        P2(idx2) = T.PkAmp(idx2);
    end

    summaryTbl = table(x, DC, Rs, Rm, P1, P2, ...
        'VariableNames', {xLabel, 'DC_pA', 'Rs_MOhm', 'Rm_MOhm', ...
                          'PkAmp_Pulse1_pA', 'PkAmp_Pulse2_pA'});

    % ---------- P2 / P1 ratio ----------
    if ismember('PkAmp', vars)

        % Filename column for grouping
        if ismember('Filename', vars)
            fn = string(T.Filename);
        elseif ismember('Filename_Ext', vars)
            fn = string(T.Filename_Ext);
        else
            fn = string((1:n)');
        end

        [G,~,gid] = unique(fn);
        nSweeps = numel(G);
        ratioX = nan(nSweeps,1);
        ratioY = nan(nSweeps,1);

        for k = 1:nSweeps
            rows = gid == k;
            idx1 = rows & pul == 1;
            idx2 = rows & pul == 2;
            if any(idx1) && any(idx2)
                p1 = T.PkAmp(find(idx1,1));
                p2 = T.PkAmp(find(idx2,1));
                if p1 ~= 0 && ~isnan(p1) && ~isnan(p2)
                    ratioY(k) = p2 / p1;
                    ratioX(k) = mean(x(rows));   % time for this sweep
                end
            end
        end

        valid = ~isnan(ratioY);
        ratioTbl = table(ratioX(valid), ratioY(valid), ...
            'VariableNames', {xLabel, 'P2_over_P1'});
    end

    % ---------- Drug marker table ----------
    if isempty(app.DrugMarkers)
        drugTbl = table();
        return;
    end

    % choose filename column again
    fnVar = '';
    if ismember('Filename_Ext', vars)
        fnVar = 'Filename_Ext';
    elseif ismember('Filename', vars)
        fnVar = 'Filename';
    else
        drugTbl = table();
        return;
    end
    fnCol = string(T.(fnVar));

    names   = strings(0,1);
    types   = strings(0,1);  % 'on' or 'off'
    files   = strings(0,1);
    xvals   = zeros(0,1);

    for k = 1:numel(app.DrugMarkers)
        dm = app.DrugMarkers(k);

        if dm.onFile ~= ""
            idx = find(fnCol == dm.onFile, 1);
            if ~isempty(idx)
                names(end+1,1) = dm.name;
                types(end+1,1) = "on";
                files(end+1,1) = dm.onFile;
                xvals(end+1,1) = x(idx);
            end
        end
        if dm.offFile ~= ""
            idx = find(fnCol == dm.offFile, 1);
            if ~isempty(idx)
                names(end+1,1) = dm.name;
                types(end+1,1) = "off";
                files(end+1,1) = dm.offFile;
                xvals(end+1,1) = x(idx);
            end
        end
    end

    if ~isempty(names)
        drugTbl = table(names, types, files, xvals, ...
            'VariableNames', {'Drug', 'MarkerType', 'Filename', xLabel});
    else
        drugTbl = table();
    end
end

function updateP0FileLists(app)
       % ----- Base list: all sweeps from directory (.P0 + .P1) -----
    if isempty(app.P0Files)
        allNames = {};
    else
        allNames = {app.P0Files.name};
    end

    % ----- Extension filter from dropdown -----
    extFiltered = allNames;
    if ~isempty(allNames) && isprop(app, 'SweepExtDropDown')
        switch app.SweepExtDropDown.Value
            case 'P0 only'
                mask = endsWith(lower(allNames), '.p0');
                extFiltered = allNames(mask);
            case 'P1 only'
                mask = endsWith(lower(allNames), '.p1');
                extFiltered = allNames(mask);
            otherwise % 'All'
                extFiltered = allNames;
        end
    end

    items = extFiltered;   % will possibly be further filtered by workbook

    % ----- Workbook-based filtering (if enabled) -----
    if app.FilterP0ByWorkbook && ~isempty(app.CurrentWorkbook) && ~isempty(items)
        T    = app.CurrentWorkbook;
        vars = T.Properties.VariableNames;

        wbNames = {};
        if ismember('Filename_Ext', vars)
            wbNames = cellstr(T.Filename_Ext);         % e.g. '5D020012.P0'
        elseif ismember('Filename', vars)
            wbNames = strcat(cellstr(T.Filename), '.P0');  % fallback assumption
        end

        if ~isempty(wbNames)
            wbNames = unique(strtrim(wbNames));

            % case-insensitive intersection
            itemsLower = lower(items);
            wbLower    = lower(wbNames);
            keep       = ismember(itemsLower, wbLower);

            items = items(keep);

            % safety: if nothing matched, fall back to extFiltered
            if isempty(items)
                items = extFiltered;
            end
        end
    end

       % ----- Update main sweep listbox -----
    prevVal = '';
    if ~isempty(app.FileListBox.Items) && ~isempty(app.FileListBox.Value)
        prevVal = app.FileListBox.Value;
    end

    if isempty(items)
        % No matching files (e.g. .P1 only in a P0-only folder)
        app.FileListBox.Items = {};
        app.FileListBox.Value = {};   % <-- must be {} when Items is {}
    else
        app.FileListBox.Items = items;

        if any(strcmp(prevVal, items))
            app.FileListBox.Value = prevVal;
        else
            app.FileListBox.Value = items{1};
        end
    end

    % ----- Update Drug On/Off listboxes -----
    drugBoxes = {app.Drug1OnListBox, app.Drug1OffListBox, ...
                 app.Drug2OnListBox, app.Drug2OffListBox, ...
                 app.Drug3OnListBox, app.Drug3OffListBox, ...
                 app.Drug4OnListBox, app.Drug4OffListBox};

    for i = 1:numel(drugBoxes)
        lb = drugBoxes{i};

        prev = '';
        if ~isempty(lb.Items) && ~isempty(lb.Value)
            prev = lb.Value;
        end

        % For drug listboxes, we always keep at least the blank entry
        lb.Items = [{''}, items];

        if any(strcmp(prev, lb.Items))
            lb.Value = prev;
        else
            lb.Value = '';
        end
    end

% ----- Update Sweep Reanalysis listbox (if it exists) -----
if isprop(app, 'SelectSweepListBox')
    prevReVal = '';
    if ~isempty(app.SelectSweepListBox.Items) && ~isempty(app.SelectSweepListBox.Value)
        prevReVal = app.SelectSweepListBox.Value;
    end

    app.SelectSweepListBox.Items = items;

    if isempty(items)
        app.SelectSweepListBox.Value = {};
    else
        if any(strcmp(prevReVal, items))
            app.SelectSweepListBox.Value = prevReVal;
        else
            app.SelectSweepListBox.Value = items{1};
        end
    end
end


end


function u = getTimePlotYUnit(app)
    u = strtrim(app.TimePlotUnit);   % e.g. "1 pA" or "1 mV" or ""

    if startsWith(u, "1")
        parts = split(u);
        if numel(parts) >= 2
            u = parts(2);           % "pA" or "mV"
        end
    end

    if u == ""
        u = "pA";                   % safe fallback
    end
end

function updateTimePlotParamDropdowns(app)
    vars = app.TimePlotParamVars;
    if isempty(vars)
        items = {'(none)'};
    else
        items = [{'(none)'}, vars{:}];
    end

    % Set items
    app.DCParamDropDown.Items     = items;
    app.RsParamDropDown.Items     = items;
    app.RmParamDropDown.Items     = items;
    app.PkAmpParamDropDown.Items  = items;

    % Helper to choose default if available
    function val = pickDefault(name)
        if any(strcmp(name, vars))
            val = name;
        else
            val = '(none)';
        end
    end

    % Reasonable defaults for intracellular workbooks
    app.DCParamDropDown.Value    = pickDefault('DC');
    app.PkAmpParamDropDown.Value = pickDefault('PkAmp');
    app.RsParamDropDown.Value    = pickDefault('Rs_Mohm');
    app.RmParamDropDown.Value    = pickDefault('Rm_Mohm');

    % For extracellular, you’ll probably see:
    %   DCParam -> '(none)' or maybe 'DC'
    %   PkAmpParam -> 'PkAmp'
    %   Rs/RmParam -> '(none)'
end


function plotTimeParamOnAxis(app, ax, paramName, x, xLabel, pul, T)
    % paramName: name of column or '(none)'

    vars = T.Properties.VariableNames;

    if paramName == "(none)" || ~ismember(paramName, vars)
        cla(ax);
        return;
    end

    % Special cases: DC, Rs, Rm, PkAmp
    % 1) DC (pulse 1 only, uses workbook unit)
    if paramName == "DC"
        idx = (pul == 1) & ~isnan(T.DC);
        plot(ax, x(idx), T.DC(idx));
        xlabel(ax, xLabel);
        ylabel(ax, sprintf('DC (%s)', app.getTimePlotYUnit()));
        title(ax, 'DC (pulse 1)');
        app.drawDrugMarkersOnAxes(ax, x, T);
        return;
    end

    % 2) Rs and Rm (pulse 1 only, MOhm units)
    if paramName == "Rs_Mohm"
        idx = (pul == 1) & ~isnan(T.Rs_Mohm);
        plot(ax, x(idx), T.Rs_Mohm(idx));
        xlabel(ax, xLabel);
        ylabel(ax, 'Rs (M\Omega)');
        title(ax, 'Rs (pulse 1)');
        app.drawDrugMarkersOnAxes(ax, x, T);
        return;
    end

    if paramName == "Rm_Mohm"
        idx = (pul == 1) & ~isnan(T.Rm_Mohm);
        plot(ax, x(idx), T.Rm_Mohm(idx));
        xlabel(ax, xLabel);
        ylabel(ax, 'Rm (M\Omega)');
        title(ax, 'Rm (pulse 1)');
        app.drawDrugMarkersOnAxes(ax, x, T);
        return;
    end

    % 3) PkAmp: show pulse 1 & 2 separately, with legend
    if paramName == "PkAmp"
        cla(ax); hold(ax,'on');

        idx1 = pul == 1 & ~isnan(T.PkAmp);
        idx2 = pul == 2 & ~isnan(T.PkAmp);

        if any(idx1)
            plot(ax, x(idx1), T.PkAmp(idx1), 'DisplayName', 'Pulse 1');
        end
        if any(idx2)
            plot(ax, x(idx2), T.PkAmp(idx2), 'DisplayName', 'Pulse 2');
        end

        hold(ax,'off');
        xlabel(ax, xLabel);
        ylabel(ax, sprintf('PkAmp (%s)', app.getTimePlotYUnit()));
        title(ax, 'Peak amplitude');
        if any(idx1) || any(idx2)
            legend(ax, 'Location', 'best');
        end
        app.drawDrugMarkersOnAxes(ax, x, T);
        return;
    end

    % 4) Generic parameter: plot as-is
    y = T.(paramName);
    plot(ax, x, y);
    xlabel(ax, xLabel);

    if contains(paramName, 'Mohm', 'IgnoreCase', true)
        ylabel(ax, sprintf('%s (M\Omega)', paramName));
    else
        ylabel(ax, sprintf('%s (%s)', paramName, app.getTimePlotYUnit()));
    end

    title(ax, paramName);
    app.drawDrugMarkersOnAxes(ax, x, T);
end

function updateReanalysisCursors(app)
   if isempty(app.ReanX)
        return;
    end
    app.refreshReanalysisCursors();
    %hold(ax, 'off');
end


function computeReanalysisMetrics(app)
       if isempty(app.ReanX) || isempty(app.ReanY)
        app.LastReanPeak1     = NaN;
        app.LastReanPeak1Time = NaN;
        app.LastReanPeak2     = NaN;
        app.LastReanPeak2Time = NaN;
        app.LastReanRm        = NaN;
        app.PeakResultLabel.Text = 'Peak: —';
        app.RmResultLabel.Text   = 'Rm: —';
        return;
    end

    t = app.ReanX;
    y = app.ReanY;

    % ---- helper for peak in a window ----
    function [pVal, pTime] = peakInWindow(w0, w1, base0, base1)
        pVal  = NaN;  pTime = NaN;
        if isnan(w0) || isnan(w1), return; end

        b0 = min(base0, base1);
        b1 = max(base0, base1);
        w0 = min(w0, w1);
        w1 = max(w0, w1);

        baseMask = (t >= b0) & (t <= b1);
        winMask  = (t >= w0) & (t <= w1);

        if ~any(baseMask) || ~any(winMask), return; end

        baseline = mean(y(baseMask));
        segT = t(winMask);
        segY = y(winMask);

        [~, idxMax] = max(segY);
        [~, idxMin] = min(segY);
        dMax = segY(idxMax) - baseline;
        dMin = segY(idxMin) - baseline;
        if abs(dMax) >= abs(dMin)
            idxPeak = idxMax;
        else
            idxPeak = idxMin;
        end

        % N-pt average around extremum
        if isprop(app,'PeakAvgPointsEditField') && ~isempty(app.PeakAvgPointsEditField.Value)
            N = max(1, round(app.PeakAvgPointsEditField.Value));
        else
            N = 5;
        end
        N = min(N, numel(segY));
        halfWin = floor(N/2);
        i0 = max(1, idxPeak-halfWin);
        i1 = min(numel(segY), idxPeak+halfWin);

        pSeg   = segY(i0:i1);
        pVal   = mean(pSeg) - baseline;
        pTime  = mean(segT(i0:i1));
    end

    % ----- Peak1 -----
    [app.LastReanPeak1, app.LastReanPeak1Time] = ...
        peakInWindow(app.Peak1Start, app.Peak1End, app.BaselineStart, app.BaselineEnd);

    % (update label however you want, e.g. show peak1)
    % if ~isnan(app.LastReanPeak1)
    %     app.PeakResultLabel.Text = sprintf('Peak: %.2f %s at %.1f ms', ...
    %         app.LastReanPeak1, app.ReanUnit, app.LastReanPeak1Time);
    % else
    %     app.PeakResultLabel.Text = 'Peak: —';
    % end

    % ----- Peak2 (optional) -----
    [app.LastReanPeak2, app.LastReanPeak2Time] = ...
        peakInWindow(app.Peak2Start, app.Peak2End, app.BaselineStart, app.BaselineEnd);

p1 = app.LastReanPeak1;
p2 = app.LastReanPeak2;
t1 = app.LastReanPeak1Time;
t2 = app.LastReanPeak2Time;

if ~isnan(p1) && ~isnan(p2)
    app.PeakResultLabel.Text = sprintf( ...
        'Peak1: %.2f %s @ %.1f ms   |   Peak2: %.2f %s @ %.1f ms', ...
        p1, app.ReanUnit, t1, ...
        p2, app.ReanUnit, t2);
elseif ~isnan(p1)
    app.PeakResultLabel.Text = sprintf( ...
        'Peak1: %.2f %s @ %.1f ms', p1, app.ReanUnit, t1);
elseif ~isnan(p2)
    app.PeakResultLabel.Text = sprintf( ...
        'Peak2: %.2f %s @ %.1f ms', p2, app.ReanUnit, t2);
else
    app.PeakResultLabel.Text = 'Peaks: —';
end


    % ----- Rm -----
    app.LastReanRm = NaN;

    rb0 = min(app.RmBaselineStart, app.RmBaselineEnd);
    rb1 = max(app.RmBaselineStart, app.RmBaselineEnd);
    rw0 = min(app.RmWindowStart,   app.RmWindowEnd);
    rw1 = max(app.RmWindowStart,   app.RmWindowEnd);
    rBaseMask = (t >= rb0) & (t <= rb1);
    rWinMask  = (t >= rw0) & (t <= rw1);

    if any(rBaseMask) && any(rWinMask) && ...
       isfield(app.ReanMeta,'RsRmPulseAmp_mV') && ~isnan(app.ReanMeta.RsRmPulseAmp_mV)

        rBaseline = mean(y(rBaseMask));
        rSegY     = y(rWinMask);
        Vstep_mV  = app.ReanMeta.RsRmPulseAmp_mV;
        Imean_pA  = mean(rSegY) - rBaseline;

        if Imean_pA ~= 0
            app.LastReanRm = abs(Vstep_mV / Imean_pA) * 1e3;   % MΩ
            app.RmResultLabel.Text = sprintf('Rm: %.1f MΩ', app.LastReanRm);
        else
            app.RmResultLabel.Text = 'Rm: ∞';
        end
    else
        app.RmResultLabel.Text = 'Rm: —';
    end

end



   function initReanalysisTable(app)
    app.ReanalysisTable.ColumnName = { ...
        'Sweep', ...
        'Peak1', 'Peak1 Time (ms)', ...
        'Peak2', 'Peak2 Time (ms)', ...
        'Rm (MΩ)'};
    app.ReanalysisTable.Data = {};   % start empty
end


function addReanalysisToTable(app)

    sweepName = app.SelectSweepListBox.Value;

    if isempty(sweepName)
        uialert(app.QuickPlotterUIFigure, 'No sweep selected.', 'Error');
        return;
    end

    % Make sure metrics are up to date
    app.computeReanalysisMetrics();

    % Create a new row: Sweep | Peak1 | t1 | Peak2 | t2 | Rm
    newRow = { ...
        sweepName, ...
        app.LastReanPeak1,     ...
        app.LastReanPeak1Time, ...
        app.LastReanPeak2,     ...
        app.LastReanPeak2Time, ...
        app.LastReanRm         ...
        };

    % Append row to table
    oldData = app.ReanalysisTable.Data;
    if isempty(oldData)
        app.ReanalysisTable.Data = newRow;
    else
        app.ReanalysisTable.Data = [oldData; newRow];
    end

    %app.updateCurrentWorkbookWithReanalysis();

end


function refreshReanalysisCursors(app)
   ax = app.UIAxes;
    if isempty(ax) || ~isvalid(ax)
        return;
    end

    % ---------- delete old lines safely ----------
    % Baseline
    ln = app.BaselineLines;
    if ~isempty(ln)
        delete(ln(isgraphics(ln)));
    end

    % Peak 1
    ln = app.Peak1Lines;
    if ~isempty(ln)
        delete(ln(isgraphics(ln)));
    end

    % Peak 2
    ln = app.Peak2Lines;
    if ~isempty(ln)
        delete(ln(isgraphics(ln)));
    end

    % Rm baseline
    ln = app.RmBaselineLines;
    if ~isempty(ln)
        delete(ln(isgraphics(ln)));
    end

    % Rm window
    ln = app.RmWindowLines;
    if ~isempty(ln)
        delete(ln(isgraphics(ln)));
    end

    hold(ax, 'on');

    %========== BASELINE (shared) ==========
    if ~isnan(app.BaselineStart) && ~isnan(app.BaselineEnd)
        app.BaselineLines(1) = xline(ax, app.BaselineStart, '--', ...
            'Color',[0.3 0.3 0.3]);
        app.BaselineLines(2) = xline(ax, app.BaselineEnd,   '--', ...
            'Color',[0.3 0.3 0.3]);
    else
        app.BaselineLines = gobjects(1,2);
    end

    %========== PEAK 1 WINDOW ==========
    if ~isnan(app.Peak1Start) && ~isnan(app.Peak1End)
        app.Peak1Lines(1) = xline(ax, app.Peak1Start, '-.', ...
            'Color',[0.3 0.3 0.3]);
        app.Peak1Lines(2) = xline(ax, app.Peak1End,   '-.', ...
            'Color',[0.3 0.3 0.3]);
    else
        app.Peak1Lines = gobjects(1,2);
    end

    %========== PEAK 2 WINDOW (optional) ==========
    if ~isnan(app.Peak2Start) && ~isnan(app.Peak2End)
        app.Peak2Lines(1) = xline(ax, app.Peak2Start, '-.', ...
            'Color',[0.6 0.2 0.2]);   % different colour
        app.Peak2Lines(2) = xline(ax, app.Peak2End,   '-.', ...
            'Color',[0.6 0.2 0.2]);
    else
        app.Peak2Lines = gobjects(1,2);
    end

    %========== Rm BASELINE ==========
    if ~isnan(app.RmBaselineStart) && ~isnan(app.RmBaselineEnd)
        app.RmBaselineLines(1) = xline(ax, app.RmBaselineStart, '--', ...
            'Color',[0 0.3 0.8]);
        app.RmBaselineLines(2) = xline(ax, app.RmBaselineEnd,   '--', ...
            'Color',[0 0.3 0.8]);
    else
        app.RmBaselineLines = gobjects(1,2);
    end

    %========== Rm WINDOW ==========
    if ~isnan(app.RmWindowStart) && ~isnan(app.RmWindowEnd)
        app.RmWindowLines(1) = xline(ax, app.RmWindowStart, '-.', ...
            'Color',[0 0.3 0.8]);
        app.RmWindowLines(2) = xline(ax, app.RmWindowEnd,   '-.', ...
            'Color',[0 0.3 0.8]);
    else
        app.RmWindowLines = gobjects(1,2);
    end

    hold(ax, 'off');
end

function nUpdated = applyReanalysisRmToWorkbook(app)
    % Update Rm_Mohm in CurrentWorkbook for the currently selected sweep
    % using app.LastReanRm. Returns 0 or 1.

    nUpdated = 0;

    if isempty(app.CurrentWorkbook) || isempty(app.SelectSweepListBox.Value)
        disp('Rm DEBUG: No workbook or no sweep selected.');
        return;
    end

    T    = app.CurrentWorkbook;
    vars = T.Properties.VariableNames;

    if ~ismember('Rm_Mohm', vars)
        disp('Rm DEBUG: No Rm_Mohm column in workbook.');
        return;
    end
    if isnan(app.LastReanRm)
        disp('Rm DEBUG: LastReanRm is NaN; nothing to apply.');
        return;
    end

    % ---- Prefer Filename_Ext, else Filename ----
    if ismember('Filename_Ext', vars)
        fnVar = 'Filename_Ext';
    elseif ismember('Filename', vars)
        fnVar = 'Filename';
    else
        disp('Rm DEBUG: No Filename or Filename_Ext column.');
        return;
    end

    sweepName = string(app.SelectSweepListBox.Value);

    % DEBUG: show what we are matching against
    disp('Rm DEBUG: ------------------------------');
    disp(['Rm DEBUG: Using filename column: ', fnVar]);
    disp(['Rm DEBUG: Target sweepName: ', char(sweepName)]);
    disp('Rm DEBUG: First few workbook names:');
    disp(string(T.(fnVar)(1:min(5,height(T)))).');

    rowsForSweep = find(string(T.(fnVar)) == sweepName);
    disp(['Rm DEBUG: rowsForSweep = ', mat2str(rowsForSweep.')]);

    if isempty(rowsForSweep)
        disp('Rm DEBUG: No matching rows for this sweep.');
        return;
    end

    % Pulse column (Pul_) if available
    pulVar = vars(startsWith(vars,'Pul'));
    if ~isempty(pulVar)
        pul = T.(pulVar{1});
    else
        pul = ones(height(T),1);
    end

    rowRm = rowsForSweep(pul(rowsForSweep) == 1);
    if isempty(rowRm)
        rowRm = rowsForSweep(1);
    end

    % DEBUG: before/after values
    oldVal = T.Rm_Mohm(rowRm(1));
    newVal = app.LastReanRm;
    fprintf('Rm DEBUG: rowRm(1) = %d\n', rowRm(1));
    fprintf('Rm DEBUG: Before Rm_Mohm = %.3f\n', oldVal);
    fprintf('Rm DEBUG: New    Rm_Mohm = %.3f\n', newVal);

    T.Rm_Mohm(rowRm(1)) = newVal;
    nUpdated = 1;

    app.CurrentWorkbook      = T;
    app.WorkbookUITable.Data = T;
    app.updateTimePlots();

    disp('Rm DEBUG: Rm update applied and plots refreshed.');
end


function nUpdated = applyReanalysisPeaksToWorkbook(app)
    % Update PkAmp for current sweep using LastReanPeak1/2.
    % Returns 0, 1, or 2.

    nUpdated = 0;

    if isempty(app.CurrentWorkbook) || isempty(app.SelectSweepListBox.Value)
        disp('Pk DEBUG: No workbook or no sweep selected.');
        return;
    end

    T    = app.CurrentWorkbook;
    vars = T.Properties.VariableNames;

    if ~ismember('PkAmp', vars)
        disp('Pk DEBUG: No PkAmp column in workbook.');
        return;
    end

    if ~isprop(app,'LastReanPeak1')
        disp('Pk DEBUG: LastReanPeak1 property not found.');
        return;
    end

    if ismember('Filename_Ext', vars)
        fnVar = 'Filename_Ext';
    elseif ismember('Filename', vars)
        fnVar = 'Filename';
    else
        disp('Pk DEBUG: No Filename or Filename_Ext column.');
        return;
    end

    sweepName    = string(app.SelectSweepListBox.Value);

    disp('Pk DEBUG: ------------------------------');
    disp(['Pk DEBUG: Using filename column: ', fnVar]);
    disp(['Pk DEBUG: Target sweepName: ', char(sweepName)]);
    disp('Pk DEBUG: First few workbook names:');
    disp(string(T.(fnVar)(1:min(5,height(T)))).');

    rowsForSweep = find(string(T.(fnVar)) == sweepName);
    disp(['Pk DEBUG: rowsForSweep = ', mat2str(rowsForSweep.')]);

    if isempty(rowsForSweep)
        disp('Pk DEBUG: No matching rows for this sweep.');
        return;
    end

    pulVar = vars(startsWith(vars,'Pul'));
    if ~isempty(pulVar)
        pul = T.(pulVar{1});
    else
        pul = ones(height(T),1);
    end

    p1 = app.LastReanPeak1;
    p2 = app.LastReanPeak2;

    % Peak1 → Pul == 1
    if ~isnan(p1)
        row1 = rowsForSweep(pul(rowsForSweep) == 1);
        if isempty(row1) && numel(rowsForSweep) == 1
            row1 = rowsForSweep(1);
        end

        if ~isempty(row1)
            oldVal = T.PkAmp(row1(1));
            fprintf('Pk DEBUG: Peak1 row = %d, old = %.3f, new = %.3f\n', ...
                row1(1), oldVal, p1);
            T.PkAmp(row1(1)) = p1;
            nUpdated = nUpdated + 1;
        end
    else
        disp('Pk DEBUG: LastReanPeak1 is NaN; not updating Peak1.');
    end

    % Peak2 → Pul == 2
    if ~isnan(p2)
        row2 = rowsForSweep(pul(rowsForSweep) == 2);
        if ~isempty(row2)
            oldVal = T.PkAmp(row2(1));
            fprintf('Pk DEBUG: Peak2 row = %d, old = %.3f, new = %.3f\n', ...
                row2(1), oldVal, p2);
            T.PkAmp(row2(1)) = p2;
            nUpdated = nUpdated + 1;
        else
            disp('Pk DEBUG: No Pul==2 row for this sweep; Peak2 not applied.');
        end
    else
        disp('Pk DEBUG: LastReanPeak2 is NaN; not updating Peak2.');
    end

    if nUpdated > 0
        app.CurrentWorkbook      = T;
        app.WorkbookUITable.Data = T;
        app.updateTimePlots();
        disp('Pk DEBUG: Peaks update applied and plots refreshed.');
    end
end





    
    end
    

    % Callbacks that handle component events
    methods (Access = private)

        % Button pushed function: ChooseDirectoryButton
        function ChooseDirectoryButtonPushed(app, event)
            dirname = uigetdir();
    if isequal(dirname, 0)
        return;  % user cancelled
    end

    app.P0Dir = string(dirname);
    app.DirectoryEditField.Value = dirname;

     % Get list of *.P0 and *.P1 files
    filesP0 = dir(fullfile(dirname, '*.P0'));
    filesP1 = dir(fullfile(dirname, '*.P1'));
    files   = [filesP0; filesP1];

    app.P0Files = files;   % now holds P0 and P1

    % Reset markers and workbook-related state
    app.DrugMarkers = struct([]);
    app.CurrentWorkbook = table();

    % Reset workbook list
    xlsFiles = [dir(fullfile(app.P0Dir, '*.xls')); ...
                dir(fullfile(app.P0Dir, '*.xlsx'))];
    app.WorkbookFiles = xlsFiles;
    if isempty(xlsFiles)
        app.WorkbookListBox.Items = {};
    else
        app.WorkbookListBox.Items = {xlsFiles.name};
    end

    % Reset filter flag and checkbox
    app.FilterP0ByWorkbook = false;
    app.FilterP0ByWorkbookCheckBox.Value = false;

    % Let ONE function manage all the P0-related listboxes
    app.updateP0FileLists();

    % Clear plots / averages when directory changes
    cla(app.SweepFileAxes);
    cla(app.AVGFileAxes);
    app.LastSweepX = [];
    app.LastSweepY = [];
    app.AvgX1 = []; app.AvgY1 = []; app.AvgCount1 = 0;
    app.AvgX2 = []; app.AvgY2 = []; app.AvgCount2 = 0;
    app.AvgX3 = []; app.AvgY3 = []; app.AvgCount3 = 0;
    app.AvgX4 = []; app.AvgY4 = []; app.AvgCount4 = 0;

    %initialize table for reanalysis tab
    app.initReanalysisTable();

        end

        % Value changed function: FileListBox
        function FileListBoxValueChanged(app, event)
            value = app.FileListBox.Value;
               if isempty(app.P0Dir)
        return;
    end

    % Selected file name
    fname = app.FileListBox.Value;
    fullpath = fullfile(app.P0Dir, fname);

 % NEW: 4 outputs + meta
    [t_ms, y_AD0, unitStr, Vm_mV, meta] = app.readP0File(fullpath);

    app.LastSweepX    = t_ms;
    app.LastSweepY    = y_AD0;
    app.LastSweepUnit = unitStr;

    % Plot single sweep
    plot(app.SweepFileAxes, t_ms, y_AD0);
    xlabel(app.SweepFileAxes, 'Time (ms)');
    ylabel(app.SweepFileAxes, sprintf('AD0 (%s)', app.LastSweepUnit));
    title(app.SweepFileAxes, sprintf('%s, Vm = %.1f mV', fname, Vm_mV));
    
    % Store for averaging
    app.LastSweepX = t_ms;
    app.LastSweepY = y_AD0;
        end

        % Button pushed function: AddtoAVG1Button
        function AddtoAVG1ButtonPushed(app, event)
            
            app.addCurrentSweepToAverage(1);
            
          
        end

        % Button pushed function: SaveAVGtoTXTfileButton
        function SaveAVGtoTXTfileButtonPushed(app, event)
          
    if app.AvgCount1 + app.AvgCount2 + app.AvgCount3 + app.AvgCount4 == 0
        uialert(app.QuickPlotterUIFigure, 'No averages available to save.', 'Save Error');
        return;
    end

    totalSweeps = app.AvgCount1 + app.AvgCount2 + app.AvgCount3 + app.AvgCount4;
    defaultName = sprintf('AVGs_%dsweeps.txt', totalSweeps);

    [file, path] = uiputfile({'*.txt','Text File (*.txt)'}, ...
                             'Save All Averages As...', ...
                             fullfile(app.P0Dir, defaultName));
    if isequal(file,0)
        return;
    end

    fullpath = fullfile(path, file);
    writeAllAveragesTextFile(app, fullpath);

    uialert(app.QuickPlotterUIFigure, 'All averages saved successfully!', 'Success');
        end

        % Button pushed function: AddtoAVG2Button
        function AddtoAVG2ButtonPushed(app, event)
            app.addCurrentSweepToAverage(2);
        end

        % Button pushed function: AddtoAVG3Button
        function AddtoAVG3ButtonPushed(app, event)
            app.addCurrentSweepToAverage(3);
        end

        % Button pushed function: AddtoAVG4Button
        function AddtoAVG4ButtonPushed(app, event)
            app.addCurrentSweepToAverage(4);
        end

        % Button pushed function: ClearAVG1Button
        function ClearAVG1ButtonPushed(app, event)
            app.clearAverage(1);
        end

        % Button pushed function: ClearAVG2Button
        function ClearAVG2ButtonPushed(app, event)
            app.clearAverage(2);
        end

        % Button pushed function: ClearAVG3Button
        function ClearAVG3ButtonPushed(app, event)
            app.clearAverage(3);
        end

        % Button pushed function: ClearAVG4Button
        function ClearAVG4ButtonPushed(app, event)
            app.clearAverage(4)
        end

        % Value changed function: WorkbookListBox
        function WorkbookListBoxValueChanged(app, event)
            value = app.WorkbookListBox.Value;
         if isempty(app.P0Dir)
        return;
    end
    if isempty(app.WorkbookListBox.Value)
        return;
    end

    fname    = app.WorkbookListBox.Value;
    fullpath = fullfile(app.P0Dir, fname);

    % Read Excel as a table; let MATLAB make nice variable names
    T = readtable(fullpath);   % <- no PreserveVariableNames

    app.CurrentWorkbook = T;
    app.updateP0FileLists();

    vars = T.Properties.VariableNames;
    if ~ismember("Include", vars)
        T.Include = true(height(T), 1);   % default: all sweeps included
    end

% --- parameter columns: everything after 'Unit' ---
unitIdx = find(strcmp(vars, 'Unit'), 1);
if isempty(unitIdx)
    app.TimePlotParamVars = {};
else
    app.TimePlotParamVars = vars(unitIdx+1:end);   % e.g. {'DC','PkAmp','Rs_Mohm','Rm_Mohm'}
end

% --- capture Unit string (e.g. '1 pA', '1 mV') ---
if ismember('Unit', vars)
    ucol = string(T.Unit);
    ucol = ucol(ucol ~= "");
    if ~isempty(ucol)
        app.TimePlotUnit = ucol(1);
    else
        app.TimePlotUnit = "";
    end
else
    app.TimePlotUnit = "";
end

% update dropdowns based on this workbook
app.updateTimePlotParamDropdowns();
   
% Preview in UITable (optionally limit rows)
    maxRows = 200;
    if height(T) > maxRows
        app.WorkbookUITable.Data = T(1:maxRows, :);
    else
        app.WorkbookUITable.Data = T;
    end
    app.WorkbookUITable.ColumnName = T.Properties.VariableNames;
    app.WorkbookUITable.RowName    = {};

    vars = T.Properties.VariableNames;

% Make only Include column editable (checkbox-style)
    nCols = numel(app.WorkbookUITable.ColumnName);
    editable = false(1, nCols);
    idxInc = find(strcmp(app.WorkbookUITable.ColumnName, "Include"));
    if ~isempty(idxInc)
        editable(idxInc) = true;
    end
    app.WorkbookUITable.ColumnEditable = editable;


if ismember('Unit', vars)
    ucol = string(T.Unit);
    ucol = ucol(ucol ~= "");   % drop empty
    if ~isempty(ucol)
        app.TimePlotUnit = ucol(1);   % e.g. "1 pA" or "1 mV"
    else
        app.TimePlotUnit = "";
    end
else
    app.TimePlotUnit = "";
end
    % Update the time plots
    app.updateTimePlots();
        end

        % Value changed function: TimeBaseDropDown
        function TimeBaseDropDownValueChanged(app, event)
            value = app.TimeBaseDropDown.Value;
            app.updateTimePlots();
        end

        % Value changed function: Drug1EditField
        function Drug1EditFieldValueChanged(app, event)
            value = app.Drug1EditField.Value;
            app.syncDrugMarkersFromUI();
            app.updateTimePlots();
        end

        % Value changed function: Drug1OnListBox
        function Drug1OnListBoxValueChanged(app, event)
            value = app.Drug1OnListBox.Value;
            app.syncDrugMarkersFromUI();
    app.updateTimePlots();
        end

        % Value changed function: Drug1OffListBox
        function Drug1OffListBoxValueChanged(app, event)
            value = app.Drug1OffListBox.Value;
            app.syncDrugMarkersFromUI();
    app.updateTimePlots();
        end

        % Value changed function: Drug2EditField
        function Drug2EditFieldValueChanged(app, event)
            value = app.Drug2EditField.Value;
            app.syncDrugMarkersFromUI();
    app.updateTimePlots();
        end

        % Value changed function: Drug2OnListBox
        function Drug2OnListBoxValueChanged(app, event)
            value = app.Drug2OnListBox.Value;
            app.syncDrugMarkersFromUI();
    app.updateTimePlots();
        end

        % Value changed function: Drug2OffListBox
        function Drug2OffListBoxValueChanged(app, event)
            value = app.Drug2OffListBox.Value;
            app.syncDrugMarkersFromUI();
    app.updateTimePlots();
        end

        % Value changed function: Drug3EditField
        function Drug3EditFieldValueChanged(app, event)
            value = app.Drug3EditField.Value;
            app.syncDrugMarkersFromUI();
    app.updateTimePlots();
        end

        % Value changed function: Drug3OnListBox
        function Drug3OnListBoxValueChanged(app, event)
            value = app.Drug3OnListBox.Value;
            app.syncDrugMarkersFromUI();
    app.updateTimePlots();
        end

        % Value changed function: Drug3OffListBox
        function Drug3OffListBoxValueChanged(app, event)
            value = app.Drug3OffListBox.Value;
            app.syncDrugMarkersFromUI();
    app.updateTimePlots();
        end

        % Value changed function: Drug4EditField
        function Drug4EditFieldValueChanged(app, event)
            value = app.Drug4EditField.Value;
            app.syncDrugMarkersFromUI();
    app.updateTimePlots();
        end

        % Value changed function: Drug4OnListBox
        function Drug4OnListBoxValueChanged(app, event)
            value = app.Drug4OnListBox.Value;
            app.syncDrugMarkersFromUI();
    app.updateTimePlots();
        end

        % Value changed function: Drug4OffListBox
        function Drug4OffListBoxValueChanged(app, event)
            value = app.Drug4OffListBox.Value;
            app.syncDrugMarkersFromUI();
    app.updateTimePlots();
        end

        % Button pushed function: SavePlotstoWorkbookButton
        function SavePlotstoWorkbookButtonPushed(app, event)
             if isempty(app.CurrentWorkbook) || height(app.CurrentWorkbook) == 0
        uialert(app.UIFigure, 'No workbook loaded on Time Plot tab.', ...
            'Save Error');
        return;
    end

    % Build tables from current plots
    [summaryTbl, ratioTbl, drugTbl, xLabel] = app.buildTimePlotExportTables();

    [file, path] = uiputfile({'*.xlsx','Excel Workbook (*.xlsx)'}, ...
                             'Save time plots to Excel', ...
                             fullfile(app.P0Dir, 'TimePlots.xlsx'));

    if isequal(file,0)
        return;  % user cancelled
    end

    fullpath = fullfile(path, file);

    % Sheet 1: original workbook (as seen in table)
    try
        writetable(app.CurrentWorkbook, fullpath, 'Sheet', 'RawWorkbook');
    catch ME
        uialert(app.QuickPlotterUIFigure, sprintf('Error writing Excel file:\n%s', ME.message), ...
            'Save Error');
        return;
    end

    % Sheet 2: time-series used for DC, Rs, Rm, PkAmp plots
    if ~isempty(summaryTbl)
        writetable(summaryTbl, fullpath, 'Sheet', 'TimeSeries');
    end

    % Sheet 3: P2/P1 ratio
    if ~isempty(ratioTbl)
        writetable(ratioTbl, fullpath, 'Sheet', 'P2_P1_Ratio');
    end

    % Sheet 4: drug marker positions
    if ~isempty(drugTbl)
        writetable(drugTbl, fullpath, 'Sheet', 'DrugMarkers');
    end

    uialert(app.QuickPlotterUIFigure, 'Graphs data exported to Excel successfully.', ...
        'Export complete');
        end

        % Value changed function: FilterP0ByWorkbookCheckBox
        function FilterP0ByWorkbookCheckBoxValueChanged(app, event)
            %value = app.FilterP0ByWorkbookCheckBox.Value;
            %function FilterP0ByWorkbookCheckBoxValueChanged(app, event)
    app.FilterP0ByWorkbook = app.FilterP0ByWorkbookCheckBox.Value;
    app.updateP0FileLists();
%end

        end

        % Value changed function: DCParamDropDown
        function DCParamDropDownValueChanged(app, event)
            value = app.DCParamDropDown.Value;
            app.updateTimePlots();
        end

        % Value changed function: RsParamDropDown
        function RsParamDropDownValueChanged(app, event)
            value = app.RsParamDropDown.Value;
            app.updateTimePlots();
        end

        % Value changed function: RmParamDropDown
        function RmParamDropDownValueChanged(app, event)
            value = app.RmParamDropDown.Value;
            app.updateTimePlots();
        end

        % Value changed function: PkAmpParamDropDown
        function PkAmpParamDropDownValueChanged(app, event)
            value = app.PkAmpParamDropDown.Value;
            app.updateTimePlots();
        end

        % Button pushed function: ExportTimePlotsButton
        function ExportTimePlotsButtonPushed(app, event)
            % Ask user where to save
    folder = uigetdir(app.P0Dir, "Select folder to save plots");
    if isequal(folder,0)
        return;
    end

    % List of axes to export
    axesList = {
        app.DCAxes,     'DC';
        app.RsAxes,     'Rs';
        app.RmAxes,     'Rm';
        app.PkAmpAxes,  'PkAmp'
    };

    if isprop(app, 'RatioAxes')
        axesList(end+1,:) = {app.RatioAxes, 'Ratio'}; %#ok<AGROW>
    end

    for k = 1:size(axesList,1)
        ax = axesList{k,1};
        label = axesList{k,2};

        % Create filename
        fname = fullfile(folder, sprintf('%s_plot.png', label));

        % Export the axes as an image
        exportgraphics(ax, fname, 'Resolution', 300);
    end

    uialert(app.QuickPlotterUIFigure, 'Plots exported successfully!', 'Done');
        end

        % Value changed function: SweepExtDropDown
        function SweepExtDropDownValueChanged(app, event)
            value = app.SweepExtDropDown.Value;
            % Just recompute the list of files whenever the extension filter changes
            app.updateP0FileLists();
        end

        % Value changed function: StimFilterDropDown
        function StimFilterDropDownValueChanged(app, event)
            value = app.StimFilterDropDown.Value;
            app.updateTimePlots();
        end

        % Value changed function: SelectSweepListBox
        function SelectSweepListBoxValueChanged(app, event)
            value = app.SelectSweepListBox.Value;
           if isempty(app.P0Dir) || isempty(app.SelectSweepListBox.Items)
        return;
    end

    fname = app.SelectSweepListBox.Value;
    if isempty(fname)
        return;
    end

    fullpath = fullfile(app.P0Dir, fname);

    % Read sweep
    [t_ms, y_AD0, unitStr, Vm_mV, meta] = app.readP0File(fullpath);

    app.ReanX    = t_ms;
    app.ReanY    = y_AD0;
    app.ReanUnit = unitStr;
    app.ReanMeta = meta;

    % Plot
    ax = app.UIAxes;
    cla(ax);
    plot(ax, t_ms, y_AD0);
    xlabel(ax, 'Time (ms)');
    ylabel(ax, sprintf('AD0 (%s)', unitStr));
    title(ax, sprintf('%s (Vm = %.1f mV)', fname, Vm_mV));

    % ----- Initialize windows only ONCE -----
    if ~app.ReanWindowsInitialized
        % Peak cursors – set whatever defaults you like here:
        app.BaselineStart = 0;
        app.BaselineEnd   = min(5,  t_ms(end)/4);
        app.Peak1Start  = min(10, t_ms(end)/2);
        app.Peak1End    = min(20, t_ms(end));

        % Rm cursors – optional defaults, only once:
        % app.RmBaselineStart = 300;
        % app.RmBaselineEnd   = 350;
        % app.RmWindowStart   = 500;
        % app.RmWindowEnd     = 550;

        app.ReanWindowsInitialized = true;
    end

    % Sync edit fields to current properties (no resetting!)
    app.BaselineStartEditField.Value = app.BaselineStart;
    app.BaselineEndEditField.Value   = app.BaselineEnd;
    app.Peak1StartEditField.Value  = app.Peak1Start;
    app.Peak1EndEditField.Value    = app.Peak1End;

    app.RmBaselineStartEditField.Value = app.RmBaselineStart;
    app.RmBaselineEndEditField.Value   = app.RmBaselineEnd;
    app.RmWindowStartEditField.Value   = app.RmWindowStart;
    app.RmWindowEndEditField.Value     = app.RmWindowEnd;

    % Draw cursors and compute metrics
    app.refreshReanalysisCursors();   % or app.updateReanalysisCursors(), but just one of them
    app.computeReanalysisMetrics();

        end

        % Value changed function: BaselineStartEditField
        function BaselineStartEditFieldValueChanged(app, event)
            value = app.BaselineStartEditField.Value;
            app.BaselineStart = app.BaselineStartEditField.Value;
            app.refreshReanalysisCursors();
            app.updateReanalysisCursors();
            app.computeReanalysisMetrics();
        end

        % Value changed function: BaselineEndEditField
        function BaselineEndEditFieldValueChanged(app, event)
            value = app.BaselineEndEditField.Value;
            app.BaselineEnd = app.BaselineEndEditField.Value;
            app.refreshReanalysisCursors();
            app.updateReanalysisCursors();
            app.computeReanalysisMetrics();
        end

        % Value changed function: Peak1StartEditField
        function Peak1StartEditFieldValueChanged(app, event)
            value = app.Peak1StartEditField.Value;
            app.Peak1Start = app.Peak1StartEditField.Value;
            app.refreshReanalysisCursors();
            app.updateReanalysisCursors();
            app.computeReanalysisMetrics();
        end

        % Value changed function: Peak1EndEditField
        function Peak1EndEditFieldValueChanged(app, event)
            value = app.Peak1EndEditField.Value;
            app.Peak1End = app.Peak1EndEditField.Value;
            app.refreshReanalysisCursors();
            app.updateReanalysisCursors();
            app.computeReanalysisMetrics();
        end

        % Button pushed function: RecomputeButton
        function RecomputeButtonPushed(app, event)
            app.updateReanalysisCursors();
            app.computeReanalysisMetrics();
        end

        % Callback function
        function AddPeaktoTableButtonPushed(app, event)
              sweepName = app.SelectSweepListBox.Value;
    if isempty(sweepName)
        uialert(app.QuickPlotterUIFigure, 'No sweep selected.', 'Error');
        return;
    end

    % Use stored computed values
    peak     = app.LastReanPeak;
    peakTime = app.LastReanPeakTime;

    % Rm is not being added here
    Rm = NaN;

    newRow = {sweepName, peak, peakTime, Rm};

    oldData = app.ReanalysisTable.Data;
    if isempty(oldData)
        app.ReanalysisTable.Data = newRow;
    else
        app.ReanalysisTable.Data = [oldData; newRow];
    end

    % ---- Auto-advance to next sweep ----
    items = app.SelectSweepListBox.Items;
    cur   = app.SelectSweepListBox.Value;

    idx = find(strcmp(items, cur), 1);
    if ~isempty(idx) && idx < numel(items)
        app.SelectSweepListBox.Value = items{idx+1};
        % reload plot / recompute for new sweep
        app.SelectSweepListBoxValueChanged([]);
    end

        end

        % Callback function
        function AddRmtoTableButtonPushed(app, event)
                sweepName = app.SelectSweepListBox.Value;
    if isempty(sweepName)
        uialert(app.QuickPlotterUIFigure, 'No sweep selected.', 'Error');
        return;
    end

    % Grab the value from computeReanalysisMetrics
    Rm = app.LastReanRm;

    % Not adding peak here
    peak     = NaN;
    peakTime = NaN;

    newRow = {sweepName, peak, peakTime, Rm};

    oldData = app.ReanalysisTable.Data;
    if isempty(oldData)
        app.ReanalysisTable.Data = newRow;
    else
        app.ReanalysisTable.Data = [oldData; newRow];
    end

    % ---- Auto-advance to next sweep ----
items = app.SelectSweepListBox.Items;
cur   = app.SelectSweepListBox.Value;

idx = find(strcmp(items, cur), 1);
if ~isempty(idx) && idx < numel(items)
    app.SelectSweepListBox.Value = items{idx+1};
    app.SelectSweepListBoxValueChanged([]);
end

        end

        % Value changed function: RmBaselineStartEditField
        function RmBaselineStartEditFieldValueChanged(app, event)
            value = app.RmBaselineStartEditField.Value;
            app.RmBaselineStart = app.RmBaselineStartEditField.Value;
           app.refreshReanalysisCursors();
           app.computeReanalysisMetrics();   % optional but handy
        end

        % Value changed function: RmBaselineEndEditField
        function RmBaselineEndEditFieldValueChanged(app, event)
            value = app.RmBaselineEndEditField.Value;
           app.RmBaselineEnd = app.RmBaselineEndEditField.Value;
    app.refreshReanalysisCursors();
    app.computeReanalysisMetrics();
        end

        % Value changed function: RmWindowStartEditField
        function RmWindowStartEditFieldValueChanged(app, event)
            value = app.RmWindowStartEditField.Value;
            app.RmWindowStart = app.RmWindowStartEditField.Value;
    app.refreshReanalysisCursors();
    app.computeReanalysisMetrics();
        end

        % Value changed function: RmWindowEndEditField
        function RmWindowEndEditFieldValueChanged(app, event)
            value = app.RmWindowEndEditField.Value;
           app.RmWindowEnd = app.RmWindowEndEditField.Value;
           app.refreshReanalysisCursors();
          app.computeReanalysisMetrics();
        end

        % Button pushed function: AddAlltoTableButton
        function AddAlltoTableButtonPushed(app, event)
            sweepName = app.SelectSweepListBox.Value;
    if isempty(sweepName)
        uialert(app.QuickPlotterUIFigure, 'No sweep selected.', 'Error');
        return;
    end

    app.addReanalysisToTable();

    % ---- Auto-advance to next sweep ----
    %items = app.SelectSweepListBox.Items;
    %cur   = app.SelectSweepListBox.Value;

    %idx = find(strcmp(items, cur), 1);
    %if ~isempty(idx) && idx < numel(items)
       % app.SelectSweepListBox.Value = items{idx+1};
        %app.SelectSweepListBoxValueChanged([]);   % reload next sweep, keep cursors
    %end
        end

        % Button pushed function: SaveTabletoExcelButton
        function SaveTabletoExcelButtonPushed(app, event)
 
    %--- 1. Get data from the UI table ---
    data = app.ReanalysisTable.Data;
    if isempty(data)
        uialert(app.QuickPlotterUIFigure, ...
            'Reanalysis table is empty – nothing to save.', ...
            'Save Error');
        return;
    end

    colNames = app.ReanalysisTable.ColumnName;   % cell array of headers

    % Make valid MATLAB variable names for writetable
    varNames = matlab.lang.makeValidName(colNames, 'ReplacementStyle', 'delete');

    % Convert cell array -> table
    T = cell2table(data, 'VariableNames', varNames);

    %--- 2. Ask user where to save ---
    defaultName = 'ReanalysisTable.xlsx';
    if ~isempty(app.P0Dir)
        defaultFull = fullfile(app.P0Dir, defaultName);
    else
        defaultFull = defaultName;
    end

    [fname, fpath] = uiputfile({'*.xlsx','Excel Workbook (*.xlsx)'}, ...
                               'Save reanalysis table as...', ...
                               defaultFull);
    if isequal(fname,0)
        % user cancelled
        return;
    end

    fullpath = fullfile(fpath, fname);

    %--- 3. Write main table to sheet "Reanalysis" ---
    writetable(T, fullpath, 'Sheet', 'Reanalysis');

    %--- 4. Save window settings on 2nd sheet ---
    try
        Settings = table( ...
            app.BaselineStart,  app.BaselineEnd,  ...
            app.Peak1Start,     app.Peak1End,     ...
            app.Peak2Start,     app.Peak2End,     ...
            app.RmBaselineStart,app.RmBaselineEnd,...
            app.RmWindowStart,  app.RmWindowEnd,  ...
            'VariableNames', { ...
                'PeakBaselineStart_ms', 'PeakBaselineEnd_ms', ...
                'Peak1WindowStart_ms',  'Peak1WindowEnd_ms',  ...
                'Peak2WindowStart_ms',  'Peak2WindowEnd_ms',  ...
                'RmBaselineStart_ms',   'RmBaselineEnd_ms',   ...
                'RmWindowStart_ms',     'RmWindowEnd_ms'} );

        writetable(Settings, fullpath, ...
                   'Sheet', 'Windows', 'WriteMode', 'overwritesheet');
    catch
        % If something goes wrong here, we still have the main table saved
        warning('Could not write window settings sheet, but main table was saved.');
    end

    %--- 5. Confirm to user ---
    uialert(app.QuickPlotterUIFigure, ...
        sprintf('Reanalysis table saved to:\n%s', fullpath), ...
        'Saved');

        end

        % Value changed function: Peak2StartEditField
        function Peak2StartEditFieldValueChanged(app, event)
            value = app.Peak2StartEditField.Value;
            app.Peak2Start = app.Peak2StartEditField.Value;
            app.refreshReanalysisCursors();
            app.updateReanalysisCursors();
            app.computeReanalysisMetrics();
        end

        % Value changed function: Peak2EndEditField
        function Peak2EndEditFieldValueChanged(app, event)
            value = app.Peak2EndEditField.Value;
            app.Peak2End = app.Peak2EndEditField.Value;
            app.refreshReanalysisCursors();
            app.updateReanalysisCursors();
            app.computeReanalysisMetrics();
        end

        % Button pushed function: ApplyReanalysisRmButton
        function ApplyReanalysisRmButtonPushed(app, event)
   n = app.applyReanalysisRmToWorkbook();
    msg = sprintf('Applied Rm reanalysis to %d row(s).', n);
    uialert(app.QuickPlotterUIFigure, msg, 'Rm updated');
            
            %           % --- basic checks ---
    % if isempty(app.CurrentWorkbook) || height(app.CurrentWorkbook) == 0
    %     uialert(app.QuickPlotterUIFigure, ...
    %         'No workbook is loaded.', 'Apply Rm');
    %     return;
    % end
    % 
    % data = app.ReanalysisTable.Data;
    % if isempty(data)
    %     uialert(app.QuickPlotterUIFigure, ...
    %         'Reanalysis table is empty – nothing to apply.', 'Apply Rm');
    %     return;
    % end
    % 
    % T = app.CurrentWorkbook;
    % vars = T.Properties.VariableNames;
    % 
    % % We assume the workbook has a Filename_Ext and Rm_Mohm column
    % if ~ismember('Rm_Mohm', vars)
    %     uialert(app.QuickPlotterUIFigure, ...
    %         'Current workbook has no Rm_Mohm column.', 'Apply Rm');
    %     return;
    % end
    % 
    % % Try to use Filename_Ext if it exists; otherwise build it from Filename & Ext
    % if ismember('Filename_Ext', vars)
    %     sweepNames = string(T.Filename_Ext);
    % elseif ismember('Filename', vars) && ismember('Ext', vars)
    %     sweepNames = string(T.Filename) + "." + string(T.Ext);
    % else
    %     uialert(app.QuickPlotterUIFigure, ...
    %         ['Could not find a sweep name column (Filename_Ext or ', ...
    %          'Filename/Ext) to match reanalysis entries.'], ...
    %          'Apply Rm');
    %     return;
    % end
    % 
    % % The reanalysis table structure:
    % %  Column 1: Sweep name  (e.g. '5D020025.P0')
    % %  Column 4: Rm (MOhm)
    % sweepCol = 1;
    % rmCol    = 4;
    % 
    % nRows = size(data,1);
    % nUpdated = 0;
    % 
    % for k = 1:nRows
    %     sweepName = string(data{k, sweepCol});
    %     newRm     = data{k, rmCol};
    % 
    %     % Only apply if numeric and not empty
    %     if isempty(newRm) || ~isnumeric(newRm) || isnan(newRm)
    %         continue;
    %     end
    % 
    %     % Find matching rows in workbook
    %     mask = (sweepNames == sweepName);
    % 
    %     if ~any(mask)
    %         % no matching sweep – silently skip or log if you like
    %         continue;
    %     end
    % 
    %     % For Rm we typically only care about pulse 1 row, if Pul_ exists
    %     if ismember('Pul_', vars)
    %         rowIdx = find(mask & T.Pul_ == 1, 1);
    %         if isempty(rowIdx)
    %             rowIdx = find(mask, 1);  % fallback: first matching row
    %         end
    %     else
    %         rowIdx = find(mask, 1);
    %     end
    % 
    %     if ~isempty(rowIdx)
    %         T.Rm_Mohm(rowIdx) = newRm;
    %         nUpdated = nUpdated + 1;
    %     end
    % end
    % 
    % % Commit updated table back to the app
    % app.CurrentWorkbook = T;
    % 
    % % Refresh the workbook UITable preview (if you’re showing it)
    % maxRows = 200;
    % if height(T) > maxRows
    %     app.WorkbookUITable.Data = T(1:maxRows,:);
    % else
    %     app.WorkbookUITable.Data = T;
    % end
    % 
    % % Redraw time-plots with the new Rm values
    % app.updateTimePlots();
    % n = app.updateCurrentWorkbookWithReanalysis();
    % 
    % % User feedback
    % uialert(app.QuickPlotterUIFigure, ...
    %     sprintf('Applied reanalysis Rm to %d sweep(s).', n), ...
    %     'Apply Rm');
        end

        % Button pushed function: ApplyReanalysisPeaksButton
        function ApplyReanalysisPeaksButtonPushed(app, event)
         n = app.applyReanalysisPeaksToWorkbook();      
    msg = sprintf('Applied peak reanalysis to %d row(s).', n);
    uialert(app.QuickPlotterUIFigure, msg, 'Peaks updated');
        end

        % Callback function
        function ApplyAllReanalysistoWorkbookandTimePlotButtonPushed(app, event)
           
        end

        % Button pushed function: ToggleIncludeButton
        function ToggleIncludeButtonPushed(app, event)
           if isempty(app.CurrentWorkbook)
        return;
    end

    % Figure out which sweep name to use
    sweepName = "";

    % 1) Prefer the Reanalysis listbox if it has a value
    if isprop(app,'SelectSweepListBox') && ~isempty(app.SelectSweepListBox.Value)
        sweepName = string(app.SelectSweepListBox.Value);

    % 2) Otherwise fall back to the Sweep Select tab listbox
    elseif isprop(app,'FileListBox') && ~isempty(app.FileListBox.Value)
        sweepName = string(app.FileListBox.Value);
    end

    % If we still don't have a sweep name, bail out
    if strlength(sweepName) == 0
        return;
    end

    T    = app.CurrentWorkbook;
    vars = T.Properties.VariableNames;

    % Ensure Include column exists
    if ~ismember("Include", vars)
        T.Include = true(height(T), 1);
        vars = T.Properties.VariableNames;
    end

    % Find a filename column (e.g. 'Filename_Ext' or 'Filename')
    fnVar = vars(startsWith(vars, "Filename"));
    if isempty(fnVar)
        uialert(app.QuickPlotterUIFigure, ...
                'Could not find Filename column in workbook.', ...
                'Toggle Include');
        return;
    end
    fnVar = fnVar{1};

    rows = string(T.(fnVar)) == sweepName;

    if ~any(rows)
        uialert(app.QuickPlotterUIFigure, ...
                sprintf('Sweep %s not found in current workbook.', sweepName), ...
                'Toggle Include');
        return;
    end

    % Toggle Include flag for this sweep
    currentVal    = T.Include(rows);
    newVal        = ~currentVal;
    T.Include(rows) = newVal;

    app.CurrentWorkbook      = T;
    app.WorkbookUITable.Data = T;
    app.updateTimePlots();

    % Optional: status message
    if all(newVal)
        msg = sprintf('Sweep %s is now INCLUDED in plots.', sweepName);
    else
        msg = sprintf('Sweep %s is now EXCLUDED from plots.', sweepName);
    end
    fprintf('%s\n', msg);
        end
    end

    % Component initialization
    methods (Access = private)

        % Create UIFigure and components
        function createComponents(app)

            % Create QuickPlotterUIFigure and hide until all components are created
            app.QuickPlotterUIFigure = uifigure('Visible', 'off');
            app.QuickPlotterUIFigure.Position = [100 100 1301 974];
            app.QuickPlotterUIFigure.Name = 'QuickPlotter';

            % Create TabGroup
            app.TabGroup = uitabgroup(app.QuickPlotterUIFigure);
            app.TabGroup.Position = [14 13 1276 869];

            % Create SweepSelectTab
            app.SweepSelectTab = uitab(app.TabGroup);
            app.SweepSelectTab.Title = 'Sweep Select';
            app.SweepSelectTab.BackgroundColor = [0.9412 0.9412 0.9412];

            % Create AVGFileAxes
            app.AVGFileAxes = uiaxes(app.SweepSelectTab);
            title(app.AVGFileAxes, 'AVG File')
            xlabel(app.AVGFileAxes, 'X')
            ylabel(app.AVGFileAxes, 'Y')
            zlabel(app.AVGFileAxes, 'Z')
            app.AVGFileAxes.Position = [590 111 641 453];

            % Create SweepFileAxes
            app.SweepFileAxes = uiaxes(app.SweepSelectTab);
            title(app.SweepFileAxes, 'Sweep File')
            xlabel(app.SweepFileAxes, 'X')
            ylabel(app.SweepFileAxes, 'Y')
            zlabel(app.SweepFileAxes, 'Z')
            app.SweepFileAxes.Position = [32 154 533 410];

            % Create SaveAVGtoTXTfileButton
            app.SaveAVGtoTXTfileButton = uibutton(app.SweepSelectTab, 'push');
            app.SaveAVGtoTXTfileButton.ButtonPushedFcn = createCallbackFcn(app, @SaveAVGtoTXTfileButtonPushed, true);
            app.SaveAVGtoTXTfileButton.BackgroundColor = [0 1 0];
            app.SaveAVGtoTXTfileButton.FontSize = 14;
            app.SaveAVGtoTXTfileButton.FontWeight = 'bold';
            app.SaveAVGtoTXTfileButton.Position = [868 48 155 25];
            app.SaveAVGtoTXTfileButton.Text = 'Save AVG to .TXT file';

            % Create FileListBoxLabel
            app.FileListBoxLabel = uilabel(app.SweepSelectTab);
            app.FileListBoxLabel.HorizontalAlignment = 'right';
            app.FileListBoxLabel.Position = [73 770 25 22];
            app.FileListBoxLabel.Text = 'File';

            % Create FileListBox
            app.FileListBox = uilistbox(app.SweepSelectTab);
            app.FileListBox.ValueChangedFcn = createCallbackFcn(app, @FileListBoxValueChanged, true);
            app.FileListBox.Position = [113 584 100 210];

            % Create AddtoAVG1Button
            app.AddtoAVG1Button = uibutton(app.SweepSelectTab, 'push');
            app.AddtoAVG1Button.ButtonPushedFcn = createCallbackFcn(app, @AddtoAVG1ButtonPushed, true);
            app.AddtoAVG1Button.Position = [235 759 100 22];
            app.AddtoAVG1Button.Text = 'Add to AVG 1';

            % Create AddtoAVG2Button
            app.AddtoAVG2Button = uibutton(app.SweepSelectTab, 'push');
            app.AddtoAVG2Button.ButtonPushedFcn = createCallbackFcn(app, @AddtoAVG2ButtonPushed, true);
            app.AddtoAVG2Button.Position = [235 707 100 22];
            app.AddtoAVG2Button.Text = 'Add to AVG 2';

            % Create AddtoAVG3Button
            app.AddtoAVG3Button = uibutton(app.SweepSelectTab, 'push');
            app.AddtoAVG3Button.ButtonPushedFcn = createCallbackFcn(app, @AddtoAVG3ButtonPushed, true);
            app.AddtoAVG3Button.Position = [235 655 100 22];
            app.AddtoAVG3Button.Text = 'Add to AVG 3';

            % Create AddtoAVG4Button
            app.AddtoAVG4Button = uibutton(app.SweepSelectTab, 'push');
            app.AddtoAVG4Button.ButtonPushedFcn = createCallbackFcn(app, @AddtoAVG4ButtonPushed, true);
            app.AddtoAVG4Button.Position = [235 604 100 22];
            app.AddtoAVG4Button.Text = 'Add to AVG 4';

            % Create AVG1NameEditFieldLabel
            app.AVG1NameEditFieldLabel = uilabel(app.SweepSelectTab);
            app.AVG1NameEditFieldLabel.HorizontalAlignment = 'right';
            app.AVG1NameEditFieldLabel.Position = [446 770 71 22];
            app.AVG1NameEditFieldLabel.Text = 'AVG1 Name';

            % Create AVG1NameEditField
            app.AVG1NameEditField = uieditfield(app.SweepSelectTab, 'text');
            app.AVG1NameEditField.Position = [538 770 100 22];

            % Create AVG2NameEditFieldLabel
            app.AVG2NameEditFieldLabel = uilabel(app.SweepSelectTab);
            app.AVG2NameEditFieldLabel.HorizontalAlignment = 'right';
            app.AVG2NameEditFieldLabel.Position = [445 729 75 22];
            app.AVG2NameEditFieldLabel.Text = 'AVG 2 Name';

            % Create AVG2NameEditField
            app.AVG2NameEditField = uieditfield(app.SweepSelectTab, 'text');
            app.AVG2NameEditField.Position = [538 730 100 22];

            % Create AVG3NameEditFieldLabel
            app.AVG3NameEditFieldLabel = uilabel(app.SweepSelectTab);
            app.AVG3NameEditFieldLabel.HorizontalAlignment = 'right';
            app.AVG3NameEditFieldLabel.Position = [743 770 78 22];
            app.AVG3NameEditFieldLabel.Text = 'AVG 3 Name ';

            % Create AVG3NameEditField
            app.AVG3NameEditField = uieditfield(app.SweepSelectTab, 'text');
            app.AVG3NameEditField.Position = [838 770 100 22];

            % Create AVG4NameEditFieldLabel
            app.AVG4NameEditFieldLabel = uilabel(app.SweepSelectTab);
            app.AVG4NameEditFieldLabel.HorizontalAlignment = 'right';
            app.AVG4NameEditFieldLabel.Position = [743 729 75 22];
            app.AVG4NameEditFieldLabel.Text = 'AVG 4 Name';

            % Create AVG4NameEditField
            app.AVG4NameEditField = uieditfield(app.SweepSelectTab, 'text');
            app.AVG4NameEditField.Position = [838 729 100 22];

            % Create ClearAVG1Button
            app.ClearAVG1Button = uibutton(app.SweepSelectTab, 'push');
            app.ClearAVG1Button.ButtonPushedFcn = createCallbackFcn(app, @ClearAVG1ButtonPushed, true);
            app.ClearAVG1Button.Position = [614 591 84 22];
            app.ClearAVG1Button.Text = 'Clear AVG1';

            % Create ClearAVG2Button
            app.ClearAVG2Button = uibutton(app.SweepSelectTab, 'push');
            app.ClearAVG2Button.ButtonPushedFcn = createCallbackFcn(app, @ClearAVG2ButtonPushed, true);
            app.ClearAVG2Button.Position = [775 592 100 22];
            app.ClearAVG2Button.Text = 'Clear AVG2';

            % Create ClearAVG3Button
            app.ClearAVG3Button = uibutton(app.SweepSelectTab, 'push');
            app.ClearAVG3Button.ButtonPushedFcn = createCallbackFcn(app, @ClearAVG3ButtonPushed, true);
            app.ClearAVG3Button.Position = [946 592 100 22];
            app.ClearAVG3Button.Text = 'Clear AVG3';

            % Create ClearAVG4Button
            app.ClearAVG4Button = uibutton(app.SweepSelectTab, 'push');
            app.ClearAVG4Button.ButtonPushedFcn = createCallbackFcn(app, @ClearAVG4ButtonPushed, true);
            app.ClearAVG4Button.Position = [1119 592 100 22];
            app.ClearAVG4Button.Text = 'Clear AVG4';

            % Create SweepFileExtensionTypeDropDownLabel
            app.SweepFileExtensionTypeDropDownLabel = uilabel(app.SweepSelectTab);
            app.SweepFileExtensionTypeDropDownLabel.HorizontalAlignment = 'right';
            app.SweepFileExtensionTypeDropDownLabel.FontColor = [0.0667 0.4431 0.7451];
            app.SweepFileExtensionTypeDropDownLabel.Position = [78 812 150 22];
            app.SweepFileExtensionTypeDropDownLabel.Text = 'Sweep File Extension Type';

            % Create SweepExtDropDown
            app.SweepExtDropDown = uidropdown(app.SweepSelectTab);
            app.SweepExtDropDown.Items = {'All', 'P0 only', 'P1 only'};
            app.SweepExtDropDown.ValueChangedFcn = createCallbackFcn(app, @SweepExtDropDownValueChanged, true);
            app.SweepExtDropDown.Position = [243 812 100 22];
            app.SweepExtDropDown.Value = 'All';

            % Create TimePlotTab
            app.TimePlotTab = uitab(app.TabGroup);
            app.TimePlotTab.Title = 'Time Plot';

            % Create DCAxes
            app.DCAxes = uiaxes(app.TimePlotTab);
            title(app.DCAxes, 'Title')
            xlabel(app.DCAxes, 'X')
            ylabel(app.DCAxes, 'Y')
            zlabel(app.DCAxes, 'Z')
            app.DCAxes.Position = [21 301 409 286];

            % Create PkAmpAxes
            app.PkAmpAxes = uiaxes(app.TimePlotTab);
            title(app.PkAmpAxes, 'Title')
            xlabel(app.PkAmpAxes, 'X')
            ylabel(app.PkAmpAxes, 'Y')
            zlabel(app.PkAmpAxes, 'Z')
            app.PkAmpAxes.Position = [38 17 395 247];

            % Create RsAxes
            app.RsAxes = uiaxes(app.TimePlotTab);
            title(app.RsAxes, 'Title')
            xlabel(app.RsAxes, 'X')
            ylabel(app.RsAxes, 'Y')
            zlabel(app.RsAxes, 'Z')
            app.RsAxes.Position = [439 298 422 286];

            % Create RmAxes
            app.RmAxes = uiaxes(app.TimePlotTab);
            title(app.RmAxes, 'Title')
            xlabel(app.RmAxes, 'X')
            ylabel(app.RmAxes, 'Y')
            zlabel(app.RmAxes, 'Z')
            app.RmAxes.Position = [874 305 389 282];

            % Create RatioAxes
            app.RatioAxes = uiaxes(app.TimePlotTab);
            title(app.RatioAxes, 'Title')
            xlabel(app.RatioAxes, 'X')
            ylabel(app.RatioAxes, 'Y')
            zlabel(app.RatioAxes, 'Z')
            app.RatioAxes.Position = [542 29 342 211];

            % Create WorkbookListBoxLabel
            app.WorkbookListBoxLabel = uilabel(app.TimePlotTab);
            app.WorkbookListBoxLabel.HorizontalAlignment = 'right';
            app.WorkbookListBoxLabel.Position = [5 810 59 22];
            app.WorkbookListBoxLabel.Text = 'Workbook';

            % Create WorkbookListBox
            app.WorkbookListBox = uilistbox(app.TimePlotTab);
            app.WorkbookListBox.ValueChangedFcn = createCallbackFcn(app, @WorkbookListBoxValueChanged, true);
            app.WorkbookListBox.Position = [79 760 100 74];

            % Create WorkbookUITable
            app.WorkbookUITable = uitable(app.TimePlotTab);
            app.WorkbookUITable.ColumnName = {'Column 1'; 'Column 2'; 'Column 3'; 'Column 4'};
            app.WorkbookUITable.RowName = {};
            app.WorkbookUITable.Position = [243 625 1012 177];

            % Create TimeBaseforXaxisLabel
            app.TimeBaseforXaxisLabel = uilabel(app.TimePlotTab);
            app.TimeBaseforXaxisLabel.HorizontalAlignment = 'right';
            app.TimeBaseforXaxisLabel.Position = [5 708 115 22];
            app.TimeBaseforXaxisLabel.Text = 'Time Base for X axis';

            % Create TimeBaseDropDown
            app.TimeBaseDropDown = uidropdown(app.TimePlotTab);
            app.TimeBaseDropDown.Items = {'Index', 'Time_min', 'Time_sec'};
            app.TimeBaseDropDown.ValueChangedFcn = createCallbackFcn(app, @TimeBaseDropDownValueChanged, true);
            app.TimeBaseDropDown.Position = [135 708 100 22];
            app.TimeBaseDropDown.Value = 'Index';

            % Create SavePlotstoWorkbookButton
            app.SavePlotstoWorkbookButton = uibutton(app.TimePlotTab, 'push');
            app.SavePlotstoWorkbookButton.ButtonPushedFcn = createCallbackFcn(app, @SavePlotstoWorkbookButtonPushed, true);
            app.SavePlotstoWorkbookButton.BackgroundColor = [0 1 0];
            app.SavePlotstoWorkbookButton.FontSize = 14;
            app.SavePlotstoWorkbookButton.FontWeight = 'bold';
            app.SavePlotstoWorkbookButton.Position = [992 158 175 25];
            app.SavePlotstoWorkbookButton.Text = 'Save Plots to Workbook';

            % Create FilterP0ByWorkbookCheckBox
            app.FilterP0ByWorkbookCheckBox = uicheckbox(app.TimePlotTab);
            app.FilterP0ByWorkbookCheckBox.ValueChangedFcn = createCallbackFcn(app, @FilterP0ByWorkbookCheckBoxValueChanged, true);
            app.FilterP0ByWorkbookCheckBox.Text = 'Filter Sweeps to current workbook ';
            app.FilterP0ByWorkbookCheckBox.FontSize = 18;
            app.FilterP0ByWorkbookCheckBox.FontWeight = 'bold';
            app.FilterP0ByWorkbookCheckBox.FontColor = [1 0 0];
            app.FilterP0ByWorkbookCheckBox.Position = [583 812 327 22];

            % Create Graph1Label
            app.Graph1Label = uilabel(app.TimePlotTab);
            app.Graph1Label.HorizontalAlignment = 'right';
            app.Graph1Label.Position = [28 598 36 22];
            app.Graph1Label.Text = 'Plot 1';

            % Create DCParamDropDown
            app.DCParamDropDown = uidropdown(app.TimePlotTab);
            app.DCParamDropDown.Items = {};
            app.DCParamDropDown.ValueChangedFcn = createCallbackFcn(app, @DCParamDropDownValueChanged, true);
            app.DCParamDropDown.Position = [79 598 100 22];
            app.DCParamDropDown.Value = {};

            % Create Graph2Label
            app.Graph2Label = uilabel(app.TimePlotTab);
            app.Graph2Label.HorizontalAlignment = 'right';
            app.Graph2Label.Position = [441 594 36 22];
            app.Graph2Label.Text = 'Plot 2';

            % Create RsParamDropDown
            app.RsParamDropDown = uidropdown(app.TimePlotTab);
            app.RsParamDropDown.Items = {};
            app.RsParamDropDown.ValueChangedFcn = createCallbackFcn(app, @RsParamDropDownValueChanged, true);
            app.RsParamDropDown.Position = [492 594 100 22];
            app.RsParamDropDown.Value = {};

            % Create Graph3Label
            app.Graph3Label = uilabel(app.TimePlotTab);
            app.Graph3Label.HorizontalAlignment = 'right';
            app.Graph3Label.Position = [874 594 36 22];
            app.Graph3Label.Text = 'Plot 3';

            % Create RmParamDropDown
            app.RmParamDropDown = uidropdown(app.TimePlotTab);
            app.RmParamDropDown.Items = {};
            app.RmParamDropDown.ValueChangedFcn = createCallbackFcn(app, @RmParamDropDownValueChanged, true);
            app.RmParamDropDown.Position = [925 594 100 22];
            app.RmParamDropDown.Value = {};

            % Create Graph4Label
            app.Graph4Label = uilabel(app.TimePlotTab);
            app.Graph4Label.HorizontalAlignment = 'right';
            app.Graph4Label.Position = [28 269 36 22];
            app.Graph4Label.Text = 'Plot 4';

            % Create PkAmpParamDropDown
            app.PkAmpParamDropDown = uidropdown(app.TimePlotTab);
            app.PkAmpParamDropDown.Items = {};
            app.PkAmpParamDropDown.ValueChangedFcn = createCallbackFcn(app, @PkAmpParamDropDownValueChanged, true);
            app.PkAmpParamDropDown.Position = [79 269 100 22];
            app.PkAmpParamDropDown.Value = {};

            % Create ExportTimePlotsButton
            app.ExportTimePlotsButton = uibutton(app.TimePlotTab, 'push');
            app.ExportTimePlotsButton.ButtonPushedFcn = createCallbackFcn(app, @ExportTimePlotsButtonPushed, true);
            app.ExportTimePlotsButton.BackgroundColor = [0.9294 0.6941 0.1255];
            app.ExportTimePlotsButton.FontSize = 14;
            app.ExportTimePlotsButton.FontWeight = 'bold';
            app.ExportTimePlotsButton.Position = [1004 72 152 25];
            app.ExportTimePlotsButton.Text = 'Export Plots to .PNG';

            % Create StimFilterDropDownLabel
            app.StimFilterDropDownLabel = uilabel(app.TimePlotTab);
            app.StimFilterDropDownLabel.HorizontalAlignment = 'right';
            app.StimFilterDropDownLabel.Position = [63 666 59 22];
            app.StimFilterDropDownLabel.Text = 'Stim Filter';

            % Create StimFilterDropDown
            app.StimFilterDropDown = uidropdown(app.TimePlotTab);
            app.StimFilterDropDown.Items = {'All', 'S0', 'S1'};
            app.StimFilterDropDown.ValueChangedFcn = createCallbackFcn(app, @StimFilterDropDownValueChanged, true);
            app.StimFilterDropDown.Position = [137 666 100 22];
            app.StimFilterDropDown.Value = 'All';

            % Create DrugMarkersTab
            app.DrugMarkersTab = uitab(app.TabGroup);
            app.DrugMarkersTab.Title = 'Drug Markers';

            % Create DrugOnListBoxLabel
            app.DrugOnListBoxLabel = uilabel(app.DrugMarkersTab);
            app.DrugOnListBoxLabel.HorizontalAlignment = 'right';
            app.DrugOnListBoxLabel.Position = [24 758 50 22];
            app.DrugOnListBoxLabel.Text = 'Drug On';

            % Create Drug1OnListBox
            app.Drug1OnListBox = uilistbox(app.DrugMarkersTab);
            app.Drug1OnListBox.ValueChangedFcn = createCallbackFcn(app, @Drug1OnListBoxValueChanged, true);
            app.Drug1OnListBox.Position = [89 598 100 184];

            % Create DrugOffListBoxLabel
            app.DrugOffListBoxLabel = uilabel(app.DrugMarkersTab);
            app.DrugOffListBoxLabel.HorizontalAlignment = 'right';
            app.DrugOffListBoxLabel.Position = [201 754 50 22];
            app.DrugOffListBoxLabel.Text = 'Drug Off';

            % Create Drug1OffListBox
            app.Drug1OffListBox = uilistbox(app.DrugMarkersTab);
            app.Drug1OffListBox.ValueChangedFcn = createCallbackFcn(app, @Drug1OffListBoxValueChanged, true);
            app.Drug1OffListBox.Position = [266 704 100 74];

            % Create DrugOnListBox_2Label
            app.DrugOnListBox_2Label = uilabel(app.DrugMarkersTab);
            app.DrugOnListBox_2Label.HorizontalAlignment = 'right';
            app.DrugOnListBox_2Label.Position = [32 458 50 22];
            app.DrugOnListBox_2Label.Text = 'Drug On';

            % Create Drug2OnListBox
            app.Drug2OnListBox = uilistbox(app.DrugMarkersTab);
            app.Drug2OnListBox.ValueChangedFcn = createCallbackFcn(app, @Drug2OnListBoxValueChanged, true);
            app.Drug2OnListBox.Position = [97 298 100 184];

            % Create DrugOffListBox_2Label
            app.DrugOffListBox_2Label = uilabel(app.DrugMarkersTab);
            app.DrugOffListBox_2Label.HorizontalAlignment = 'right';
            app.DrugOffListBox_2Label.Position = [206 454 50 22];
            app.DrugOffListBox_2Label.Text = 'Drug Off';

            % Create Drug2OffListBox
            app.Drug2OffListBox = uilistbox(app.DrugMarkersTab);
            app.Drug2OffListBox.ValueChangedFcn = createCallbackFcn(app, @Drug2OffListBoxValueChanged, true);
            app.Drug2OffListBox.Position = [271 404 100 74];

            % Create DrugOnListBox_3Label
            app.DrugOnListBox_3Label = uilabel(app.DrugMarkersTab);
            app.DrugOnListBox_3Label.HorizontalAlignment = 'right';
            app.DrugOnListBox_3Label.Position = [525 754 50 22];
            app.DrugOnListBox_3Label.Text = 'Drug On';

            % Create Drug3OnListBox
            app.Drug3OnListBox = uilistbox(app.DrugMarkersTab);
            app.Drug3OnListBox.ValueChangedFcn = createCallbackFcn(app, @Drug3OnListBoxValueChanged, true);
            app.Drug3OnListBox.Position = [590 594 100 184];

            % Create DrugOffListBox_3Label
            app.DrugOffListBox_3Label = uilabel(app.DrugMarkersTab);
            app.DrugOffListBox_3Label.HorizontalAlignment = 'right';
            app.DrugOffListBox_3Label.Position = [726 752 50 22];
            app.DrugOffListBox_3Label.Text = 'Drug Off';

            % Create Drug3OffListBox
            app.Drug3OffListBox = uilistbox(app.DrugMarkersTab);
            app.Drug3OffListBox.ValueChangedFcn = createCallbackFcn(app, @Drug3OffListBoxValueChanged, true);
            app.Drug3OffListBox.Position = [791 702 100 74];

            % Create DrugOnListBox_4Label
            app.DrugOnListBox_4Label = uilabel(app.DrugMarkersTab);
            app.DrugOnListBox_4Label.HorizontalAlignment = 'right';
            app.DrugOnListBox_4Label.Position = [542 452 50 22];
            app.DrugOnListBox_4Label.Text = 'Drug On';

            % Create Drug4OnListBox
            app.Drug4OnListBox = uilistbox(app.DrugMarkersTab);
            app.Drug4OnListBox.ValueChangedFcn = createCallbackFcn(app, @Drug4OnListBoxValueChanged, true);
            app.Drug4OnListBox.Position = [607 292 100 184];

            % Create DrugOffListBox_4Label
            app.DrugOffListBox_4Label = uilabel(app.DrugMarkersTab);
            app.DrugOffListBox_4Label.HorizontalAlignment = 'right';
            app.DrugOffListBox_4Label.Position = [746 456 50 22];
            app.DrugOffListBox_4Label.Text = 'Drug Off';

            % Create Drug4OffListBox
            app.Drug4OffListBox = uilistbox(app.DrugMarkersTab);
            app.Drug4OffListBox.ValueChangedFcn = createCallbackFcn(app, @Drug4OffListBoxValueChanged, true);
            app.Drug4OffListBox.Position = [811 406 100 74];

            % Create Drug1EditFieldLabel
            app.Drug1EditFieldLabel = uilabel(app.DrugMarkersTab);
            app.Drug1EditFieldLabel.HorizontalAlignment = 'right';
            app.Drug1EditFieldLabel.FontWeight = 'bold';
            app.Drug1EditFieldLabel.Position = [123 809 43 22];
            app.Drug1EditFieldLabel.Text = 'Drug 1';

            % Create Drug1EditField
            app.Drug1EditField = uieditfield(app.DrugMarkersTab, 'text');
            app.Drug1EditField.ValueChangedFcn = createCallbackFcn(app, @Drug1EditFieldValueChanged, true);
            app.Drug1EditField.FontWeight = 'bold';
            app.Drug1EditField.Position = [181 808 144 22];

            % Create Drug2Label
            app.Drug2Label = uilabel(app.DrugMarkersTab);
            app.Drug2Label.HorizontalAlignment = 'right';
            app.Drug2Label.FontWeight = 'bold';
            app.Drug2Label.Position = [143 504 43 22];
            app.Drug2Label.Text = 'Drug 2';

            % Create Drug2EditField
            app.Drug2EditField = uieditfield(app.DrugMarkersTab, 'text');
            app.Drug2EditField.ValueChangedFcn = createCallbackFcn(app, @Drug2EditFieldValueChanged, true);
            app.Drug2EditField.FontWeight = 'bold';
            app.Drug2EditField.Position = [201 503 144 22];

            % Create DrugEditField_2Label_2
            app.DrugEditField_2Label_2 = uilabel(app.DrugMarkersTab);
            app.DrugEditField_2Label_2.HorizontalAlignment = 'right';
            app.DrugEditField_2Label_2.FontWeight = 'bold';
            app.DrugEditField_2Label_2.Position = [659 808 43 22];
            app.DrugEditField_2Label_2.Text = 'Drug 3';

            % Create Drug3EditField
            app.Drug3EditField = uieditfield(app.DrugMarkersTab, 'text');
            app.Drug3EditField.ValueChangedFcn = createCallbackFcn(app, @Drug3EditFieldValueChanged, true);
            app.Drug3EditField.FontWeight = 'bold';
            app.Drug3EditField.Position = [717 807 144 22];

            % Create DrugEditField_2Label_3
            app.DrugEditField_2Label_3 = uilabel(app.DrugMarkersTab);
            app.DrugEditField_2Label_3.HorizontalAlignment = 'right';
            app.DrugEditField_2Label_3.FontWeight = 'bold';
            app.DrugEditField_2Label_3.Position = [659 504 43 22];
            app.DrugEditField_2Label_3.Text = 'Drug 4';

            % Create Drug4EditField
            app.Drug4EditField = uieditfield(app.DrugMarkersTab, 'text');
            app.Drug4EditField.ValueChangedFcn = createCallbackFcn(app, @Drug4EditFieldValueChanged, true);
            app.Drug4EditField.FontWeight = 'bold';
            app.Drug4EditField.Position = [717 503 144 22];

            % Create SweepReanalysisTab
            app.SweepReanalysisTab = uitab(app.TabGroup);
            app.SweepReanalysisTab.Title = 'Sweep Reanalysis';

            % Create UIAxes
            app.UIAxes = uiaxes(app.SweepReanalysisTab);
            title(app.UIAxes, 'Title')
            xlabel(app.UIAxes, 'X')
            ylabel(app.UIAxes, 'Y')
            zlabel(app.UIAxes, 'Z')
            app.UIAxes.Position = [230 411 659 409];

            % Create SelectSweepListBoxLabel
            app.SelectSweepListBoxLabel = uilabel(app.SweepReanalysisTab);
            app.SelectSweepListBoxLabel.HorizontalAlignment = 'right';
            app.SelectSweepListBoxLabel.Position = [10 783 78 22];
            app.SelectSweepListBoxLabel.Text = 'Select Sweep';

            % Create SelectSweepListBox
            app.SelectSweepListBox = uilistbox(app.SweepReanalysisTab);
            app.SelectSweepListBox.ValueChangedFcn = createCallbackFcn(app, @SelectSweepListBoxValueChanged, true);
            app.SelectSweepListBox.Position = [103 573 100 234];

            % Create BaselineStartmsLabel
            app.BaselineStartmsLabel = uilabel(app.SweepReanalysisTab);
            app.BaselineStartmsLabel.HorizontalAlignment = 'right';
            app.BaselineStartmsLabel.Position = [18 326 107 22];
            app.BaselineStartmsLabel.Text = 'Baseline Start (ms)';

            % Create BaselineStartEditField
            app.BaselineStartEditField = uieditfield(app.SweepReanalysisTab, 'numeric');
            app.BaselineStartEditField.ValueChangedFcn = createCallbackFcn(app, @BaselineStartEditFieldValueChanged, true);
            app.BaselineStartEditField.Position = [140 326 100 22];
            app.BaselineStartEditField.Value = 5;

            % Create BaselineEndmsLabel
            app.BaselineEndmsLabel = uilabel(app.SweepReanalysisTab);
            app.BaselineEndmsLabel.HorizontalAlignment = 'right';
            app.BaselineEndmsLabel.Position = [21 290 103 22];
            app.BaselineEndmsLabel.Text = 'Baseline End (ms)';

            % Create BaselineEndEditField
            app.BaselineEndEditField = uieditfield(app.SweepReanalysisTab, 'numeric');
            app.BaselineEndEditField.ValueChangedFcn = createCallbackFcn(app, @BaselineEndEditFieldValueChanged, true);
            app.BaselineEndEditField.Position = [139 290 100 22];
            app.BaselineEndEditField.Value = 15;

            % Create WindowstartmsLabel
            app.WindowstartmsLabel = uilabel(app.SweepReanalysisTab);
            app.WindowstartmsLabel.HorizontalAlignment = 'right';
            app.WindowstartmsLabel.Position = [21 250 98 22];
            app.WindowstartmsLabel.Text = 'Peak 1 Start (ms)';

            % Create Peak1StartEditField
            app.Peak1StartEditField = uieditfield(app.SweepReanalysisTab, 'numeric');
            app.Peak1StartEditField.ValueChangedFcn = createCallbackFcn(app, @Peak1StartEditFieldValueChanged, true);
            app.Peak1StartEditField.Position = [134 250 100 22];

            % Create WindowEndmsLabel
            app.WindowEndmsLabel = uilabel(app.SweepReanalysisTab);
            app.WindowEndmsLabel.HorizontalAlignment = 'right';
            app.WindowEndmsLabel.Position = [21 218 94 22];
            app.WindowEndmsLabel.Text = 'Peak 1 End (ms)';

            % Create Peak1EndEditField
            app.Peak1EndEditField = uieditfield(app.SweepReanalysisTab, 'numeric');
            app.Peak1EndEditField.ValueChangedFcn = createCallbackFcn(app, @Peak1EndEditFieldValueChanged, true);
            app.Peak1EndEditField.Position = [134 218 100 22];
            app.Peak1EndEditField.Value = 10;

            % Create RecomputeButton
            app.RecomputeButton = uibutton(app.SweepReanalysisTab, 'push');
            app.RecomputeButton.ButtonPushedFcn = createCallbackFcn(app, @RecomputeButtonPushed, true);
            app.RecomputeButton.Position = [100 435 100 22];
            app.RecomputeButton.Text = 'Recompute';

            % Create PeakResultLabel
            app.PeakResultLabel = uilabel(app.SweepReanalysisTab);
            app.PeakResultLabel.FontSize = 14;
            app.PeakResultLabel.FontWeight = 'bold';
            app.PeakResultLabel.Position = [14 371 492 22];
            app.PeakResultLabel.Text = 'Peak: ';

            % Create RmResultLabel
            app.RmResultLabel = uilabel(app.SweepReanalysisTab);
            app.RmResultLabel.FontSize = 14;
            app.RmResultLabel.FontWeight = 'bold';
            app.RmResultLabel.FontColor = [0 0 1];
            app.RmResultLabel.Position = [297 331 430 22];
            app.RmResultLabel.Text = 'Rm: ';

            % Create ptsforPeakavgEditFieldLabel
            app.ptsforPeakavgEditFieldLabel = uilabel(app.SweepReanalysisTab);
            app.ptsforPeakavgEditFieldLabel.HorizontalAlignment = 'right';
            app.ptsforPeakavgEditFieldLabel.Position = [13 98 105 22];
            app.ptsforPeakavgEditFieldLabel.Text = '# pts for Peak avg ';

            % Create PeakAvgPointsEditField
            app.PeakAvgPointsEditField = uieditfield(app.SweepReanalysisTab, 'numeric');
            app.PeakAvgPointsEditField.Position = [133 98 100 22];
            app.PeakAvgPointsEditField.Value = 20;

            % Create ReanalysisTable
            app.ReanalysisTable = uitable(app.SweepReanalysisTab);
            app.ReanalysisTable.ColumnName = {'Column 1'; 'Column 2'; 'Column 3'; 'Column 4'};
            app.ReanalysisTable.RowName = {};
            app.ReanalysisTable.Position = [901 301 362 519];

            % Create RmBaselineStartmsEditFieldLabel
            app.RmBaselineStartmsEditFieldLabel = uilabel(app.SweepReanalysisTab);
            app.RmBaselineStartmsEditFieldLabel.HorizontalAlignment = 'right';
            app.RmBaselineStartmsEditFieldLabel.FontColor = [0 0 1];
            app.RmBaselineStartmsEditFieldLabel.Position = [305 297 129 22];
            app.RmBaselineStartmsEditFieldLabel.Text = 'Rm Baseline Start (ms)';

            % Create RmBaselineStartEditField
            app.RmBaselineStartEditField = uieditfield(app.SweepReanalysisTab, 'numeric');
            app.RmBaselineStartEditField.ValueChangedFcn = createCallbackFcn(app, @RmBaselineStartEditFieldValueChanged, true);
            app.RmBaselineStartEditField.Position = [449 297 100 22];
            app.RmBaselineStartEditField.Value = 350;

            % Create RmBaselineEndmsLabel
            app.RmBaselineEndmsLabel = uilabel(app.SweepReanalysisTab);
            app.RmBaselineEndmsLabel.HorizontalAlignment = 'right';
            app.RmBaselineEndmsLabel.FontColor = [0 0 1];
            app.RmBaselineEndmsLabel.Position = [305 263 125 22];
            app.RmBaselineEndmsLabel.Text = 'Rm Baseline End (ms)';

            % Create RmBaselineEndEditField
            app.RmBaselineEndEditField = uieditfield(app.SweepReanalysisTab, 'numeric');
            app.RmBaselineEndEditField.ValueChangedFcn = createCallbackFcn(app, @RmBaselineEndEditFieldValueChanged, true);
            app.RmBaselineEndEditField.Position = [449 263 100 22];
            app.RmBaselineEndEditField.Value = 380;

            % Create RmMeasureStartmsLabel
            app.RmMeasureStartmsLabel = uilabel(app.SweepReanalysisTab);
            app.RmMeasureStartmsLabel.HorizontalAlignment = 'right';
            app.RmMeasureStartmsLabel.FontColor = [0 0 1];
            app.RmMeasureStartmsLabel.Position = [304 226 130 22];
            app.RmMeasureStartmsLabel.Text = 'Rm Measure Start (ms)';

            % Create RmWindowStartEditField
            app.RmWindowStartEditField = uieditfield(app.SweepReanalysisTab, 'numeric');
            app.RmWindowStartEditField.ValueChangedFcn = createCallbackFcn(app, @RmWindowStartEditFieldValueChanged, true);
            app.RmWindowStartEditField.Position = [449 226 100 22];
            app.RmWindowStartEditField.Value = 550;

            % Create RmMeasureStartmsLabel_2
            app.RmMeasureStartmsLabel_2 = uilabel(app.SweepReanalysisTab);
            app.RmMeasureStartmsLabel_2.HorizontalAlignment = 'right';
            app.RmMeasureStartmsLabel_2.FontColor = [0 0 1];
            app.RmMeasureStartmsLabel_2.Position = [304 189 126 22];
            app.RmMeasureStartmsLabel_2.Text = 'Rm Measure End (ms)';

            % Create RmWindowEndEditField
            app.RmWindowEndEditField = uieditfield(app.SweepReanalysisTab, 'numeric');
            app.RmWindowEndEditField.ValueChangedFcn = createCallbackFcn(app, @RmWindowEndEditFieldValueChanged, true);
            app.RmWindowEndEditField.Position = [450 189 100 22];
            app.RmWindowEndEditField.Value = 570;

            % Create AddAlltoTableButton
            app.AddAlltoTableButton = uibutton(app.SweepReanalysisTab, 'push');
            app.AddAlltoTableButton.ButtonPushedFcn = createCallbackFcn(app, @AddAlltoTableButtonPushed, true);
            app.AddAlltoTableButton.BackgroundColor = [0 1 1];
            app.AddAlltoTableButton.FontSize = 14;
            app.AddAlltoTableButton.FontWeight = 'bold';
            app.AddAlltoTableButton.Position = [304 147 259 25];
            app.AddAlltoTableButton.Text = 'Add Reanalysis Parameters to Table';

            % Create SaveTabletoExcelButton
            app.SaveTabletoExcelButton = uibutton(app.SweepReanalysisTab, 'push');
            app.SaveTabletoExcelButton.ButtonPushedFcn = createCallbackFcn(app, @SaveTabletoExcelButtonPushed, true);
            app.SaveTabletoExcelButton.BackgroundColor = [0.2863 0.8588 0.251];
            app.SaveTabletoExcelButton.FontSize = 14;
            app.SaveTabletoExcelButton.FontWeight = 'bold';
            app.SaveTabletoExcelButton.FontColor = [0 0 0];
            app.SaveTabletoExcelButton.Position = [1022 239 144 25];
            app.SaveTabletoExcelButton.Text = 'Save Table to Excel';

            % Create Peak2StartmsEditFieldLabel
            app.Peak2StartmsEditFieldLabel = uilabel(app.SweepReanalysisTab);
            app.Peak2StartmsEditFieldLabel.HorizontalAlignment = 'right';
            app.Peak2StartmsEditFieldLabel.FontColor = [1 0 0];
            app.Peak2StartmsEditFieldLabel.Position = [21 180 98 22];
            app.Peak2StartmsEditFieldLabel.Text = 'Peak 2 Start (ms)';

            % Create Peak2StartEditField
            app.Peak2StartEditField = uieditfield(app.SweepReanalysisTab, 'numeric');
            app.Peak2StartEditField.ValueChangedFcn = createCallbackFcn(app, @Peak2StartEditFieldValueChanged, true);
            app.Peak2StartEditField.Position = [134 180 100 22];

            % Create Peak2EndmsEditFieldLabel
            app.Peak2EndmsEditFieldLabel = uilabel(app.SweepReanalysisTab);
            app.Peak2EndmsEditFieldLabel.HorizontalAlignment = 'right';
            app.Peak2EndmsEditFieldLabel.FontColor = [1 0 0];
            app.Peak2EndmsEditFieldLabel.Position = [21 148 94 22];
            app.Peak2EndmsEditFieldLabel.Text = 'Peak 2 End (ms)';

            % Create Peak2EndEditField
            app.Peak2EndEditField = uieditfield(app.SweepReanalysisTab, 'numeric');
            app.Peak2EndEditField.ValueChangedFcn = createCallbackFcn(app, @Peak2EndEditFieldValueChanged, true);
            app.Peak2EndEditField.FontColor = [0 0 0];
            app.Peak2EndEditField.Position = [134 148 100 22];

            % Create ApplyReanalysisRmButton
            app.ApplyReanalysisRmButton = uibutton(app.SweepReanalysisTab, 'push');
            app.ApplyReanalysisRmButton.ButtonPushedFcn = createCallbackFcn(app, @ApplyReanalysisRmButtonPushed, true);
            app.ApplyReanalysisRmButton.WordWrap = 'on';
            app.ApplyReanalysisRmButton.BackgroundColor = [0.8 0.8 0.8];
            app.ApplyReanalysisRmButton.FontSize = 14;
            app.ApplyReanalysisRmButton.FontWeight = 'bold';
            app.ApplyReanalysisRmButton.Position = [655 292 160 64];
            app.ApplyReanalysisRmButton.Text = 'Apply Rm Reanalysis  to Workbook & Time Plot';

            % Create ApplyReanalysisPeaksButton
            app.ApplyReanalysisPeaksButton = uibutton(app.SweepReanalysisTab, 'push');
            app.ApplyReanalysisPeaksButton.ButtonPushedFcn = createCallbackFcn(app, @ApplyReanalysisPeaksButtonPushed, true);
            app.ApplyReanalysisPeaksButton.WordWrap = 'on';
            app.ApplyReanalysisPeaksButton.BackgroundColor = [0.8 0.8 0.8];
            app.ApplyReanalysisPeaksButton.FontSize = 14;
            app.ApplyReanalysisPeaksButton.FontWeight = 'bold';
            app.ApplyReanalysisPeaksButton.Position = [652 182 166 94];
            app.ApplyReanalysisPeaksButton.Text = 'Apply Peak Reanalysis  to Workbook & Time Plot';

            % Create ToggleIncludeButton
            app.ToggleIncludeButton = uibutton(app.SweepReanalysisTab, 'push');
            app.ToggleIncludeButton.ButtonPushedFcn = createCallbackFcn(app, @ToggleIncludeButtonPushed, true);
            app.ToggleIncludeButton.BackgroundColor = [1 1 1];
            app.ToggleIncludeButton.FontWeight = 'bold';
            app.ToggleIncludeButton.FontColor = [1 0 0];
            app.ToggleIncludeButton.Position = [78 524 147 22];
            app.ToggleIncludeButton.Text = 'Exclude/Include Sweep';

            % Create ChooseDirectoryButton
            app.ChooseDirectoryButton = uibutton(app.QuickPlotterUIFigure, 'push');
            app.ChooseDirectoryButton.ButtonPushedFcn = createCallbackFcn(app, @ChooseDirectoryButtonPushed, true);
            app.ChooseDirectoryButton.Position = [29 907 144 35];
            app.ChooseDirectoryButton.Text = 'Choose Directory';

            % Create DirectoryEditFieldLabel
            app.DirectoryEditFieldLabel = uilabel(app.QuickPlotterUIFigure);
            app.DirectoryEditFieldLabel.HorizontalAlignment = 'right';
            app.DirectoryEditFieldLabel.Position = [210 913 60 22];
            app.DirectoryEditFieldLabel.Text = 'Directory: ';

            % Create DirectoryEditField
            app.DirectoryEditField = uieditfield(app.QuickPlotterUIFigure, 'text');
            app.DirectoryEditField.Position = [285 913 264 22];

            % Show the figure after all components are created
            app.QuickPlotterUIFigure.Visible = 'on';
        end
    end

    % App creation and deletion
    methods (Access = public)

        % Construct app
        function app = QuickPlots_v101

            % Create UIFigure and components
            createComponents(app)

            % Register the app with App Designer
            registerApp(app, app.QuickPlotterUIFigure)

            if nargout == 0
                clear app
            end
        end

        % Code that executes before app deletion
        function delete(app)

            % Delete UIFigure when app is deleted
            delete(app.QuickPlotterUIFigure)
        end
    end
end