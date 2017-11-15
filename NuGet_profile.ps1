function Format-ProjectItem( $item ) {

	$item.ProjectItems | ForEach-Object {
		Format-ProjectItem( $_ )
	}

	if ($item.Name.EndsWith('.cs')) {
		$window = $item.Open('{7651A701-06E5-11D1-8EBD-00A0C90F26EA}');
		if ($window){
			Write-Host $item.Name;
			[System.Threading.Thread]::Sleep(100);
			$window.Activate();
			$item.Document.DTE.ExecuteCommand('Edit.FormatDocument');
			$item.Document.DTE.ExecuteCommand('Edit.RemoveAndSort');
			$window.Close(1);
		}
	}
}

function Format-AllDocuments {
	Process {

		$dte.Solution.Projects | ForEach-Object {$_.ProjectItems | ForEach-Object {
			Format-ProjectItem( $_ )
		} }
	}
}
