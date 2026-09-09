function [ventPptx, reportPptx, reportPdf] = createCombinedVDPReport_RH(analysisRoot, selectedMethods)
%CREATECOMBINEDVDPREPORT_RH Merge each method's own reports into combined decks.
%RH edit: Builds two separate combined presentations from files already
% written, unmodified, by the legacy pipeline into each method's own folder -
% one slide per method that has the corresponding file. On Windows this is a
% literal slide copy via PowerPoint COM; on macOS/Linux (no COM, and no
% slide-duplicate command in PowerPoint-for-Mac's AppleScript dictionary)
% slides are rebuilt with python-pptx instead (see merge_pptx.py):
%   1) Ventilation_Analysis_All_Methods_<stamp>.pptx - merges each method's
%      Ventilation_Analysis_<date>.pptx (written by the raw VDP calculators:
%      calculate_VDP_CCHMC/calculate_LB_VDP/kmeans.VDP_calculationASB/
%      Akmeans.VDP_calculationASB).
%   2) VDP_Report_All_Methods_<stamp>.pptx + matching .pdf - merges each
%      method's <Method>VDP_Report_<date>.pptx (written by
%      ThresholdVDP_Report/LBVDP_Report/KmeansVDP_Report/AKmeansVDP_Report).

    stamp = datestr(now, 'yyyymmdd_HHMMSS');

    ventPptx = iMergeMethodFiles(analysisRoot, selectedMethods, ...
        'Ventilation_Analysis_*.pptx', ...
        sprintf('Ventilation_Analysis_All_Methods_%s.pptx', stamp), false);

    reportPptx = iMergeMethodFiles(analysisRoot, selectedMethods, ...
        '*VDP_Report*.pptx', ...
        sprintf('VDP_Report_All_Methods_%s.pptx', stamp), true);
    if isempty(reportPptx)
        reportPdf = '';
    else
        reportPdf = strrep(reportPptx, '.pptx', '.pdf');
    end
end

function outFile = iMergeMethodFiles(analysisRoot, selectedMethods, ...
        globPattern, outName, alsoSavePdf)
    outFile = fullfile(analysisRoot, outName);

    sourceFiles = {};
    for k = 1:numel(selectedMethods)
        method = selectedMethods{k};
        methodFolder = fullfile(analysisRoot, method);
        matchingFiles = dir(fullfile(methodFolder, globPattern));
        if isempty(matchingFiles)
            %RH edit: fprintf, not warning() - a leftover warning('off','all')
            % from calculate_VDP_CCHMC.m would otherwise silently swallow this.
            fprintf(2, ['No %s found for %s in %s; it will be skipped in ', ...
                 '%s.\n'], globPattern, method, methodFolder, outName);
            continue
        end
        [~, newestIdx] = max([matchingFiles.datenum]);
        sourceFiles{end+1} = fullfile(methodFolder, matchingFiles(newestIdx).name); %#ok<AGROW>
    end

    if isempty(sourceFiles)
        fprintf(2, 'No files matched %s; %s was not created.\n', globPattern, outName);
        outFile = '';
        return
    end

    if ispc
        ppt = actxserver('PowerPoint.Application');
        presentation = ppt.Presentations.Add(0); % 0 = msoFalse, no visible window
        for k = 1:numel(sourceFiles)
            presentation.Slides.InsertFromFile(sourceFiles{k}, presentation.Slides.Count);
        end
        presentation.SaveAs(outFile, 24); % 24 = ppSaveAsOpenXMLPresentation (.pptx)
        if alsoSavePdf
            pdfFile = strrep(outFile, '.pptx', '.pdf');
            if isfile(pdfFile)
                delete(pdfFile);
            end
            presentation.SaveAs(pdfFile, 32); % 32 = ppSaveAsPDF
        end
        presentation.Close();
        ppt.Quit();
        delete(ppt);
    else
        %RH: no COM/InsertFromFile on macOS/Linux, and PowerPoint-for-Mac's
        % AppleScript dictionary has no slide-duplicate command either (checked
        % its .sdef directly), so slides are rebuilt with python-pptx instead.
        pythonPath = Global.getPythonExe();
        scriptPath = fullfile(fileparts(mfilename('fullpath')), 'merge_pptx.py');
        cmd = sprintf('"%s" "%s" "%s"', pythonPath, scriptPath, outFile);
        for k = 1:numel(sourceFiles)
            cmd = [cmd, sprintf(' "%s"', sourceFiles{k})]; %#ok<AGROW>
        end
        status = system(cmd);
        if status ~= 0
            fprintf(2, 'Merging %s failed; %s was not created.\n', globPattern, outName);
            outFile = '';
            return
        end
        if alsoSavePdf
            pdfFile = strrep(outFile, '.pptx', '.pdf');
            Global.pptxToPdf(outFile, pdfFile);
        end
    end
    fprintf('Combined report saved to %s\n', outFile);
end
