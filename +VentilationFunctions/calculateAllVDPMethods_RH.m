function [Ventilation, Outputs, VDPTable, MetricsTable] = calculateAllVDPMethods_RH( ...
    Ventilation, Proton, MainInput, subjectID)
%CALCULATEALLVDPMETHODS_RH Run each selected VDP method in its own folder.
%RH edit: This entire RH-named file is a new multi-method coordinator; the
% legacy calculateAllVDPMethods.m remains unmodified.
%
% This RH-only coordinator preserves all selected results in Outputs, writes
% one wide VDP summary row and one long metrics row per method, and creates a
% combined, one-slide-per-method PowerPoint report.  Legacy calculators are
% intentionally called unchanged and their artifacts are promoted from their
% internal VDP_Analysis folder to the requested method folder.

    if nargin < 4 || isempty(subjectID)
        error('VentilationFunctions:RHSubjectIDRequired', ...
            'A non-empty subjectID is required.');
    end
    subjectID = string(subjectID);

    if ~isfield(Ventilation, 'parentPath') || isempty(Ventilation.parentPath)
        error('VentilationFunctions:RHParentPathMissing', ...
            'Ventilation.parentPath must be set before multi-method VDP analysis.');
    end
    analysisRoot = char(Ventilation.parentPath);
    if ~isfolder(analysisRoot)
        mkdir(analysisRoot);
    end

    allMethods = {'Threshold','Kmeans','AKmeans','LBmean','GLBmean', ...
        'HybridLBm','GLBpercentile'};
    %RH edit: Clean only the fixed RH method folders; no rerun/archive folders.
    iCleanMethodFolders(analysisRoot, allMethods);
    Ventilation = iEnsureAncillarySourceMethods(Ventilation);
    selectedMethods = iSelectedMethods(Ventilation);
    if isempty(selectedMethods)
        error('VentilationFunctions:RHNoMethodSelected', ...
            'No VDP method is selected.');
    end

    Outputs = struct();
    Outputs.N4 = struct();
    Outputs.N4.Image = Ventilation.Image;
    MetricsTable = iEmptyMetricsTable();
    vdpValues = nan(1, numel(allMethods));

    for k = 1:numel(selectedMethods)
        method = selectedMethods{k};
        methodFolder = fullfile(analysisRoot, method);
        if ~isfolder(methodFolder)
            mkdir(methodFolder);
        end

        methodVentilation = Ventilation;
        methodVentilation.parentPath = methodFolder;
        methodVentilation.outputpath = methodFolder;
        methodVentilation = iRunMethod(method, methodVentilation, Proton, MainInput);

        %RH edit: Write the same single-method PPTX/PDF report the legacy
        % single-run pipeline produces (e.g. LBVDP_Report_<date>.pptx/.pdf),
        % unmodified, straight into this method's own folder. Must run before
        % promotion below, while the calculator's VDP_Analysis\ subfolder
        % (histogram PNGs etc.) still exists at the legacy path these report
        % functions expect.
        iWriteLegacyReport(method, methodVentilation, Proton, MainInput);

        %RH edit: Promote legacy artifacts so the public layout is <root>/<method>/.
        iPromoteLegacyArtifacts(methodFolder);

        methodOutput = iCaptureOutput(method, methodVentilation, methodFolder);
        Outputs.N4.(method) = methodOutput;
        MetricsTable = [MetricsTable; iMetricsRow(subjectID, method, ...
            methodOutput, Ventilation)]; %#ok<AGROW>

        methodIndex = find(strcmp(allMethods, method), 1, 'first');
        vdpValues(methodIndex) = methodOutput.VDP;

        % Retain the legacy calculator's fields for compatibility.  When
        % multiple methods are selected, RHOutputs is the authoritative full
        % result collection; generic fields describe the last method run.
        Ventilation = methodVentilation;
        Ventilation.parentPath = analysisRoot;
        Ventilation.outputpath = analysisRoot;
    end

    %RH edit: Run DDI/GLRLM with the explicitly selected method's defect map.
    [Ventilation, MetricsTable] = iRunAncillaryAnalysis( ...
        Ventilation, Outputs, MetricsTable, Proton, MainInput, analysisRoot);

    %RH edit: One wide row per subject plus a long per-method metrics table.
    VDPTable = table(subjectID, vdpValues(1), vdpValues(2), vdpValues(3), ...
        vdpValues(4), vdpValues(5), vdpValues(6), vdpValues(7), ...
        'VariableNames', {'SubjectID','TH60_VDP','Kmeans_VDP','AKmeans_VDP', ...
        'LBmean_VDP','GLBmean_VDP','HybridLBm_VDP','GLBpercentile_VDP'});

    %RH edit: Write the combined workbook and CSV summaries.
    iWriteSummaryFiles(analysisRoot, VDPTable, MetricsTable);
    %RH edit: Merge each method's own report PPTX (written above by
    % iWriteLegacyReport) into one combined deck, one slide per method.
    VentilationFunctions.createCombinedVDPReport_RH(analysisRoot, selectedMethods);

    Ventilation.parentPath = analysisRoot;
    Ventilation.outputpath = analysisRoot;
    Ventilation.RHSelectedMethods = selectedMethods;
end

function selected = iSelectedMethods(Ventilation)
    selected = {};
    if iIsEnabled(Ventilation, 'ThreshAnalysis')
        selected{end+1} = 'Threshold';
    end
    if iIsEnabled(Ventilation, 'Kmeans')
        selected{end+1} = 'Kmeans';
    end
    if iIsEnabled(Ventilation, 'AKmeans')
        selected{end+1} = 'AKmeans';
    end
    if iIsEnabled(Ventilation, 'LB_Analysis') && isfield(Ventilation, 'LB_Methods')
        allowed = {'LBmean','GLBmean','HybridLBm','GLBpercentile'};
        requested = cellstr(string(Ventilation.LB_Methods(:).'));
        selected = [selected, requested(ismember(requested, allowed))];
    end
    selected = unique(selected, 'stable');
end

function Ventilation = iEnsureAncillarySourceMethods(Ventilation)
    %RH edit: A selected DDI/GLRLM source is automatically analysed first.
    if iIsEnabled(Ventilation, 'DDI2D') || iIsEnabled(Ventilation, 'DDI3D')
        source = iSelectedDefectMethod(Ventilation, 'DDIDefectMap');
        Ventilation = iEnableMethod(Ventilation, source);
    end
    if iIsEnabled(Ventilation, 'GLRLM_Analysis')
        source = iSelectedDefectMethod(Ventilation, 'GLRLMDefectMap');
        Ventilation = iEnableMethod(Ventilation, source);
    end
end

function Ventilation = iEnableMethod(Ventilation, method)
    switch method
        case 'Threshold'
            Ventilation.ThreshAnalysis = 'yes';
        case 'Kmeans'
            Ventilation.Kmeans = 'yes';
        case 'AKmeans'
            Ventilation.AKmeans = 'yes';
        otherwise
            Ventilation.LB_Analysis = 'yes';
            if ~isfield(Ventilation, 'LB_Methods') || isempty(Ventilation.LB_Methods)
                Ventilation.LB_Methods = {method};
            else
                methods = cellstr(string(Ventilation.LB_Methods(:).'));
                Ventilation.LB_Methods = unique([methods, {method}], 'stable');
            end
    end
end

function Ventilation = iRunMethod(method, Ventilation, Proton, MainInput)
    switch method
        case 'Threshold'
            Ventilation = VentilationFunctions.calculate_VDP_CCHMC( ...
                Ventilation, Proton, MainInput);
        case 'Kmeans'
            Ventilation = VentilationFunctions.kmeans.VDP_calculationASB( ...
                Ventilation, Proton, MainInput);
        case 'AKmeans'
            Ventilation = VentilationFunctions.Akmeans.VDP_calculationASB( ...
                Ventilation, Proton, MainInput);
        otherwise
            Ventilation = iConfigureLinearBinning(Ventilation, method);
            Ventilation = VentilationFunctions.calculate_LB_VDP( ...
                Ventilation, Proton, MainInput);
    end
    %RH edit: plain 'close all' only closes classic figure() plot windows.
    % 'close all force' also destroys uifigure windows and everything inside
    % them - including the App Designer window that is running this very
    % callback - which crashes PerformAnalysisButtonPushed with "Invalid or
    % deleted object" partway through a multi-method run.
    close all;
end

function iWriteLegacyReport(method, Ventilation, Proton, MainInput)
    %RH edit: Reuse the exact legacy single-method report functions, unchanged,
    % so each method folder gets its own <Method>VDP_Report_<date>.pptx/.pdf -
    % identical in content to what the original single-run pipeline produces.
    % Wrapped in try/catch so one method's report failure (e.g. PowerPoint COM
    % automation hiccup) doesn't lose the VDP results already computed for the
    % rest of the multi-method run.
    try
        switch method
            case 'Threshold'
                VentilationFunctions.ThresholdVDP_Report(Ventilation, Proton, MainInput);
            case 'Kmeans'
                VentilationFunctions.KmeansVDP_Report(Ventilation, Proton, MainInput);
            case 'AKmeans'
                VentilationFunctions.AKmeansVDP_Report(Ventilation, Proton, MainInput);
            otherwise
                Ventilation = iRefreshLBHealthyRef(Ventilation, MainInput);
                VentilationFunctions.LBVDP_Report(Ventilation, Proton, MainInput);
        end
    catch ME
        %RH edit: some legacy calculators (e.g. calculate_VDP_CCHMC.m) call
        % warning('off','all') and never turn it back on, which would
        % otherwise silently swallow a warning() here for the rest of the
        % session. fprintf to stderr is not affected by that state, so this
        % failure is never hidden from the user.
        fprintf(2, 'Could not write the %s report: %s\n', method, ME.message);
    end
    close all;
end

function Ventilation = iRefreshLBHealthyRef(Ventilation, MainInput)
    %RH edit: calculateAllVDPMethods_RH runs all four LB normalizations in one
    % pass, but Ventilation.HealthyRef.LB.* is only computed once, during
    % Ventilation_Analysis.m's shared prep step, for whichever LB method was
    % selected at that time. Recompute it here per submethod - mirrors the
    % switch-case in Ventilation_Analysis.m exactly - so each LB report shows
    % the correct healthy reference for its own normalization, not whichever
    % one happened to run first.
    Age = MainInput.Age;
    if (isstring(Age) && strlength(Age) == 0) || ...
            (ischar(Age) && isempty(Age)) || (isnumeric(Age) && isempty(Age))
        Age = 0;
    end
    switch Ventilation.LB_Normalization
        case 'LBmean'
            Ventilation.HealthyRef.LB.VDPULN = ['≤', num2str(round((-0.31238 + 0.020369 * Age + 1.6449 * 0.68247),2))];
            Ventilation.HealthyRef.LB.LVV = '6±2.3';
            Ventilation.HealthyRef.LB.HVV = '5.5±3.3';
        case {'GLBmean','HybridLBm'}
            Ventilation.HealthyRef.LB.VDPULN = ['≤', num2str(round((0.11008 + 0.051214 * Age + 1.6449 * 1.5833),2))];
            Ventilation.HealthyRef.LB.LVV = '10.6±3';
            Ventilation.HealthyRef.LB.HVV = '10±3.5';
        case 'GLBpercentile'
            Ventilation.HealthyRef.LB.VDPULN = ['≤', num2str(round((-0.65729 + 0.06534 * Age + 1.6449 * 1.874),2))];
            Ventilation.HealthyRef.LB.LVV = '11±8';
            Ventilation.HealthyRef.LB.HVV = '14±8.7';
        otherwise
            warning('Unrecognized LB_Normalization method: %s', Ventilation.LB_Normalization);
    end
end

function Ventilation = iConfigureLinearBinning(Ventilation, method)
    Ventilation.LB_Normalization = method;
    switch method
        case 'LBmean'
            Ventilation.LBThresholds = [0.33,0.66,1.00,1.33,1.66];
            Ventilation.Hdist = [0.001029,0.995768,0.244090];
        case 'GLBmean'
            Ventilation.LBThresholds = ...
                [0.495130,0.752533,1.001607,1.244833,1.483568];
            Ventilation.Hdist = [-0.771445,1.136303,0.282776];
        case 'HybridLBm'
            Ventilation.LBThresholds = [0.50,0.75,1.00,1.25,1.50];
            Ventilation.Hdist = [-0.771445,1.136303,0.282776];
        case 'GLBpercentile'
            Ventilation.LBThresholds = ...
                [0.288930,0.462393,0.622368,0.773740,0.918873];
            Ventilation.Hdist = [-1.186698,0.738826,0.198451];
        otherwise
            error('VentilationFunctions:RHUnknownLBMethod', ...
                'Unsupported linear-binning method: %s', method);
    end
end

function output = iCaptureOutput(method, Ventilation, outputFolder)
    output = struct();
    output.Method = string(method);
    output.OutputFolder = string(outputFolder);
    output.VDP = NaN;
    output.BinPercent = nan(1,6);
    output.DefectMap = [];
    output.VDPPerSliceLocal = [];
    output.VDPPerSliceGlobal = [];

    switch method
        case 'Threshold'
            output.VDP = iScalar(Ventilation.Threshold.VDP);
            output.BinPercent = iPadToSix(Ventilation.Threshold.THBins);
            %RH edit: Match legacy Ventilation_Analysis.m defect-map convention -
            % only Incomplete-only/Complete-only voxels count as "defect";
            % combined and hyperventilated codes (>2) are excluded.
            thDefect = Ventilation.Threshold.defectArray;
            thDefect(thDefect > 2) = 0;
            output.DefectMap = double(thDefect > 0);
            output.VDPPerSliceLocal = Ventilation.Threshold.vdp_per_slice_local;
            output.VDPPerSliceGlobal = Ventilation.Threshold.vdp_per_slice_global;
        case 'Kmeans'
            output.VDP = iScalar(Ventilation.KmeansVDP);
            output.BinPercent = iPadToSix(Ventilation.wholelung_VDP);
            %RH edit: Match legacy convention - the defect cluster is label 1
            % (post lung-mask shift), not label 0, which is outside the lung.
            output.DefectMap = double(Ventilation.Kmeans_segmentation == 1);
        case 'AKmeans'
            output.VDP = iScalar(Ventilation.Akmeans_VDP);
            output.BinPercent = iPadToSix(Ventilation.Akmeans_ventP_raw_cor);
            output.DefectMap = double(Ventilation.Akmeans_defect_mask > 0);
        otherwise
            output.VDP = iScalar(Ventilation.LB_VDP);
            output.BinPercent = iPadToSix(Ventilation.BinsPercent);
            output.DefectMap = double(Ventilation.VentBinMap2 == 1);
            output.VDPPerSliceLocal = Ventilation.LB_vdp_per_slice_local;
            output.VDPPerSliceGlobal = Ventilation.LB_vdp_per_slice_global;
    end
end

function iPromoteLegacyArtifacts(methodFolder)
    %RH edit: Replace same-named legacy artifacts directly; RH keeps no
    % Previous_Runs or output-conflict archive folders.
    legacyFolder = fullfile(methodFolder, 'VDP_Analysis');
    if ~isfolder(legacyFolder)
        return
    end

    % Legacy calculators usually leave MATLAB's current folder here.  Move
    % away before removing the now-empty directory on Windows.
    cd(methodFolder);

    files = dir(legacyFolder);
    files = files(~ismember({files.name}, {'.','..'}));
    if isempty(files)
        rmdir(legacyFolder);
        return
    end

    for k = 1:numel(files)
        source = fullfile(legacyFolder, files(k).name);
        destination = fullfile(methodFolder, files(k).name);
        if isfile(destination) || isfolder(destination)
            if isfolder(destination)
                rmdir(destination, 's');
            end
        end
        movefile(source, destination, 'f');
    end
    rmdir(legacyFolder);
end

function iCleanMethodFolders(analysisRoot, allMethods)
    %RH edit: Delete only the seven fixed RH method folders for a clean rerun.
    cd(analysisRoot);
    for k = 1:numel(allMethods)
        methodFolder = fullfile(analysisRoot, allMethods{k});
        if isfolder(methodFolder)
            rmdir(methodFolder, 's');
        end
    end

    %RH edit: Remove the archive folder created by earlier RH versions.
    oldConflictFolder = fullfile(analysisRoot, 'Previous_Runs_RH');
    if isfolder(oldConflictFolder)
        rmdir(oldConflictFolder, 's');
    end
end

function [Ventilation, MetricsTable] = iRunAncillaryAnalysis( ...
        Ventilation, Outputs, MetricsTable, Proton, MainInput, analysisRoot)
    %RH edit: DDI and GLRLM use the binary defect map from the chosen method.
    if iIsEnabled(Ventilation, 'DDI2D') || iIsEnabled(Ventilation, 'DDI3D')
        method = iSelectedDefectMethod(Ventilation, 'DDIDefectMap');
        Ventilation = iAssignDefectMap(Ventilation, Outputs.N4.(method), ...
            'DDI', method, analysisRoot);
        if iIsEnabled(Ventilation, 'DDI2D')
            Ventilation = VentilationFunctions.calculateDDI_2D( ...
                Ventilation, Proton, MainInput);
        end
        if iIsEnabled(Ventilation, 'DDI3D')
            Ventilation = VentilationFunctions.calculateDDI_3D( ...
                Ventilation, Proton, MainInput);
        end
        iPromoteLegacyArtifacts(char(Outputs.N4.(method).OutputFolder));
        MetricsTable = iAddDDIMetrics(MetricsTable, method, Ventilation);
    end

    if iIsEnabled(Ventilation, 'GLRLM_Analysis')
        method = iSelectedDefectMethod(Ventilation, 'GLRLMDefectMap');
        Ventilation = iAssignDefectMap(Ventilation, Outputs.N4.(method), ...
            'GLRLM', method, analysisRoot);
        Ventilation = VentilationFunctions.GLRLM_Analysis(Ventilation);
        MetricsTable = iAddGLRLMMetrics(MetricsTable, method, Ventilation);
    end
    Ventilation.parentPath = analysisRoot;
    Ventilation.outputpath = analysisRoot;
end

function Ventilation = iAssignDefectMap(Ventilation, methodOutput, purpose, method, analysisRoot)
    %RH edit: Convert the selected method's binary map to the legacy DDI/GLRLM convention.
    mask = double(Ventilation.LungMask);
    if isfield(Ventilation, 'AirwayMask')
        mask(Ventilation.AirwayMask == 1) = 0;
    end
    if isfield(Ventilation, 'VesselMask')
        mask(Ventilation.VesselMask == 1) = 0;
    end
    defectMap = mask + double(methodOutput.DefectMap > 0);
    methodFolder = char(methodOutput.OutputFolder);
    Ventilation.parentPath = methodFolder;
    Ventilation.outputpath = methodFolder;
    if strcmp(purpose, 'DDI')
        Ventilation.defectMap_forDDI = defectMap;
        Ventilation.DDIDefectMap = method;
    else
        Ventilation.defectMap_forGLRLM = defectMap;
        Ventilation.GLRLMDefectMap = method;
    end
    if ~isfolder(methodFolder)
        mkdir(methodFolder);
    end
    if ~isfolder(analysisRoot)
        mkdir(analysisRoot);
    end
end

function method = iSelectedDefectMethod(Ventilation, fieldName)
    %RH edit: RH App values name the exact method; old "Linear Binning" maps to LBmean.
    if ~isfield(Ventilation, fieldName) || isempty(Ventilation.(fieldName))
        method = 'Threshold';
        return
    end
    value = char(string(Ventilation.(fieldName)));
    allowed = {'Threshold','Kmeans','AKmeans','LBmean','GLBmean', ...
        'HybridLBm','GLBpercentile'};
    if ismember(value, allowed)
        method = value;
    elseif strcmpi(value, 'Linear Binning')
        method = 'LBmean';
    else
        warning('VentilationFunctions:RHUnknownDefectMap', ...
            'Unknown %s value "%s". Using Threshold.', fieldName, value);
        method = 'Threshold';
    end
end

function MetricsTable = iAddDDIMetrics(MetricsTable, method, Ventilation)
    %RH edit: Record DDI values in the selected method's combined metrics row.
    row = MetricsTable.Method == string(method);
    if iIsEnabled(Ventilation, 'DDI2D')
        MetricsTable.DDI2DMean(row) = iFieldScalar(Ventilation, 'DDI2D_mean');
        MetricsTable.DDI2DMax(row) = iFieldScalar(Ventilation, 'DDI2D_max');
    end
    if iIsEnabled(Ventilation, 'DDI3D')
        MetricsTable.DDI3DMean(row) = iFieldScalar(Ventilation, 'DDI3D_mean');
        MetricsTable.DDI3DMax(row) = iFieldScalar(Ventilation, 'DDI3D_max');
    end
end

function MetricsTable = iAddGLRLMMetrics(MetricsTable, method, Ventilation)
    %RH edit: Record the GLRLM summary values in the selected method's row.
    row = MetricsTable.Method == string(method);
    MetricsTable.GLRLM_SRE(row) = iFieldScalar(Ventilation, 'SRE');
    MetricsTable.GLRLM_LRE(row) = iFieldScalar(Ventilation, 'LRE');
    MetricsTable.GLRLM_RP(row) = iFieldScalar(Ventilation, 'RP');
end

function iWriteSummaryFiles(analysisRoot, VDPTable, MetricsTable)
    xlsxFile = fullfile(analysisRoot, 'All_VDP_Methods.xlsx');
    vdpCsv = fullfile(analysisRoot, 'VDP_Summary.csv');
    metricsCsv = fullfile(analysisRoot, 'Method_Metrics.csv');

    combinedVDP = iMergeBySubject(iReadTable(xlsxFile, 'VDP_Summary'), VDPTable);
    combinedMetrics = iMergeBySubject( ...
        iReadTable(xlsxFile, 'Method_Metrics'), MetricsTable);

    writetable(combinedVDP, xlsxFile, 'Sheet', 'VDP_Summary', ...
        'WriteMode', 'overwritesheet');
    writetable(combinedMetrics, xlsxFile, 'Sheet', 'Method_Metrics', ...
        'WriteMode', 'overwritesheet');
    writetable(combinedVDP, vdpCsv);
    writetable(combinedMetrics, metricsCsv);
end

function T = iReadTable(xlsxFile, sheetName)
    T = table();
    if ~isfile(xlsxFile)
        return
    end
    try
        T = readtable(xlsxFile, 'Sheet', sheetName, 'TextType', 'string');
    catch
        T = table();
    end
end

function combined = iMergeBySubject(existing, newRows)
    if isempty(existing)
        combined = newRows;
        return
    end
    expected = newRows.Properties.VariableNames;
    if ~isequal(existing.Properties.VariableNames, expected)
        warning('VentilationFunctions:RHExistingSummarySchema', ...
            ['Existing RH summary uses a different schema.  It is preserved in ', ...
             'the file archive and will be replaced by the current schema.']);
        combined = newRows;
        return
    end
    existing.SubjectID = string(existing.SubjectID);
    newRows.SubjectID = string(newRows.SubjectID);
    existing(existing.SubjectID == newRows.SubjectID(1), :) = [];
    combined = [existing; newRows];
    combined = sortrows(combined, 'SubjectID');
end

function T = iEmptyMetricsTable()
    %RH edit: Extend the combined table with DDI and GLRLM metrics.
    types = [{'string','string'}, repmat({'double'}, 1, 7), ...
        {'string','string'}, repmat({'double'}, 1, 6), {'string'}, ...
        repmat({'double'}, 1, 7)];
    T = table('Size', [0 numel(types)], 'VariableTypes', types, ...
        'VariableNames', {'SubjectID','Method','VDP','Bin1Pct','Bin2Pct', ...
        'Bin3Pct','Bin4Pct','Bin5Pct','Bin6Pct','VDPPerSliceLocal', ...
        'VDPPerSliceGlobal','SNRLung','SNRVV','CoV','VHI','Skewness', ...
        'Kurtosis','OutputFolder','DDI2DMean','DDI2DMax','DDI3DMean', ...
        'DDI3DMax','GLRLM_SRE','GLRLM_LRE','GLRLM_RP'});
end

function row = iMetricsRow(subjectID, method, output, sharedVentilation)
    %RH edit: Initialise DDI/GLRLM columns; selected method values are added later.
    b = iPadToSix(output.BinPercent);
    row = table(string(subjectID), string(method), output.VDP, b(1), b(2), ...
        b(3), b(4), b(5), b(6), ...
        string(mat2str(output.VDPPerSliceLocal)), ...
        string(mat2str(output.VDPPerSliceGlobal)), ...
        iFieldScalar(sharedVentilation, 'SNR_lung'), ...
        iFieldScalar(sharedVentilation, 'SNR_vv'), ...
        iFieldScalar(sharedVentilation, 'overallMeanCV'), ...
        iFieldScalar(sharedVentilation, 'overallVHI'), ...
        iFieldScalar(sharedVentilation, 'skewness'), ...
        iFieldScalar(sharedVentilation, 'kurtosis'), ...
        string(output.OutputFolder), nan, nan, nan, nan, nan, nan, nan, ...
        'VariableNames', {'SubjectID','Method','VDP','Bin1Pct','Bin2Pct', ...
        'Bin3Pct','Bin4Pct','Bin5Pct','Bin6Pct','VDPPerSliceLocal', ...
        'VDPPerSliceGlobal','SNRLung','SNRVV','CoV','VHI','Skewness', ...
        'Kurtosis','OutputFolder','DDI2DMean','DDI2DMax','DDI3DMean', ...
        'DDI3DMax','GLRLM_SRE','GLRLM_LRE','GLRLM_RP'});
end

function value = iFieldScalar(S, fieldName)
    value = NaN;
    if isfield(S, fieldName) && ~isempty(S.(fieldName))
        value = iScalar(S.(fieldName));
    end
end

function value = iScalar(value)
    if isempty(value)
        value = NaN;
    else
        value = double(value(1));
    end
end

function values = iPadToSix(values)
    values = double(values(:)).';
    values = [values, nan(1, max(0, 6 - numel(values)))];
    values = values(1:6);
end

function tf = iIsEnabled(S, fieldName)
    tf = false;
    if ~isfield(S, fieldName)
        return
    end
    value = S.(fieldName);
    if islogical(value) || isnumeric(value)
        tf = logical(value);
    else
        tf = strcmpi(strtrim(string(value)), "yes");
    end
end
