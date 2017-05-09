var target = Argument("target", "Default");
var deployLocal = Argument("deployLocal", "DeployLocal");

var product = Argument("product_name", "ChangeDimensionSets");
var addinName = Argument("addin_name", "BenWinding." + product + ".bundle");

var solutionPath = Argument("path_solution", "../");
var projectPath = Argument("path_project", solutionPath + "ChangeDimensionSets/");
var binPath = Argument("path_bin", projectPath + "bin/Debug/");

var addinPath = Argument("path_addin", "./" + addinName + "/");

Setup(ctx =>
{
	// Executed BEFORE the first task.
	Information("Running tasks...");
});

Teardown(ctx =>
{
	// Executed AFTER the last task.
	Information("Finished running tasks.");
});

Task("Default")
.Does(() => {
	MSBuild(solutionPath + "ChangeDimensionSets.sln");
});

Task("DeployLocal")
.IsDependentOn("Default")
.Does(() => {
	var contentsDirectory = addinPath + "Contents";
	var docsDirectory = contentsDirectory + "/Documentation";
	CreateDirectory(contentsDirectory);
	CreateDirectory(docsDirectory);
	CopyFileToDirectory(binPath + "ChangeDimensionSets.dll", contentsDirectory);
	CopyFileToDirectory(projectPath + "Autodesk.ChangeDimensionSets.Inventor.addin", contentsDirectory);
	CopyFileToDirectory(projectPath + "Documentation.txt", docsDirectory);
});


RunTarget(deployLocal);