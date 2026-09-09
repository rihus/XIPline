function pptxToPdf(pptxPath, pdfPath)
%PPTXTOPDF Export a saved .pptx to .pdf using the installed PowerPoint.
%RH: new file. Windows drives PowerPoint via ActiveX/COM (actxserver), which
% does not exist on macOS. macOS has no COM, but Microsoft PowerPoint for Mac
% can be driven via AppleScript ('osascript'), so use that instead. Linux has
% no supported PowerPoint automation, so PDF export is skipped there (pptx
% report is still saved by the caller either way).

if isfile(pdfPath)
    delete(pdfPath);
end

if ispc
    ppt = actxserver('PowerPoint.Application');
    presentation = ppt.Presentations.Open(pptxPath);
    presentation.SaveAs(pdfPath, 32);
    pause(2);  % allow time for file to be written
    presentation.Close();
    ppt.Quit();
    delete(ppt);
elseif ismac
    % RH: PowerPoint for Mac is sandboxed - opening a file in a folder it
    % hasn't touched before pops a "Grant File Access" dialog, once per
    % folder. With dozens of subject/method folders that means dozens of
    % prompts. Stage every file through one fixed folder instead, so
    % PowerPoint only ever needs that one folder granted (once, ever); plain
    % MATLAB copyfile moves data in/out and isn't sandboxed.
    stagingDir = fullfile(getenv('HOME'), 'XIPline', 'pptx_staging');
    if ~isfolder(stagingDir)
        mkdir(stagingDir);
    end
    [~, baseName] = fileparts(tempname);
    stagedPptx = fullfile(stagingDir, [baseName, '.pptx']);
    stagedPdf = fullfile(stagingDir, [baseName, '.pdf']);
    copyfile(pptxPath, stagedPptx);

    % RH: explicit timeout - the default AppleEvent timeout (~120s) can be too
    % short for large, image-heavy decks (e.g. combined multi-method reports).
    script = sprintf([ ...
        'with timeout of 600 seconds\n' ...
        '    tell application "Microsoft PowerPoint"\n' ...
        '        activate\n' ...
        '        open POSIX file "%s"\n' ...
        '        delay 1\n' ...
        '        set thePres to active presentation\n' ...
        '        save thePres in POSIX file "%s" as save as PDF\n' ...
        '        close thePres saving no\n' ...
        '    end tell\n' ...
        'end timeout'], strrep(stagedPptx,'"','\"'), strrep(stagedPdf,'"','\"'));
    scriptFile = [tempname, '.scpt'];
    fid = fopen(scriptFile, 'w');
    fprintf(fid, '%s', script);
    fclose(fid);
    [status, cmdout] = system(sprintf('osascript "%s"', scriptFile));
    delete(scriptFile);

    if status == 0 && isfile(stagedPdf)
        copyfile(stagedPdf, pdfPath);
    else
        warning('Global:pptxToPdfFailed', ...
            'PDF export via PowerPoint (AppleScript) failed: %s', cmdout);
    end
    if isfile(stagedPptx); delete(stagedPptx); end
    if isfile(stagedPdf); delete(stagedPdf); end
else
    disp('Skipping PDF export (no supported PowerPoint automation on Linux).');
end
end
