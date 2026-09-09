function pythonPath = getPythonExe()
%GETPYTHONEXE Resolve the Python interpreter configured for XIPline.
%RH: new file. Reads <XIPlineRoot>/python_path.txt - the same file
% Segmentation.PerformSegmentation reads/writes - so every Python-dependent
% feature uses one consistent interpreter. Falls back to 'python3' if unset.

if ispc
    XIPlineRoot = 'C:\XIPline';
else
    XIPlineRoot = fullfile(getenv('HOME'), 'XIPline');
end

pathFile = fullfile(XIPlineRoot, 'python_path.txt');
if isfile(pathFile)
    fid = fopen(pathFile, 'r');
    pythonPath = strtrim(fgetl(fid));
    fclose(fid);
else
    pythonPath = 'python3';
end
end
