using Microsoft.SharePoint.Client;

namespace KeenMate.FluentlySharePoint.Extensions
{
	public static class Folder
	{
		public static CSOMOperation CreateFolder(this CSOMOperation operation, string remotePath, bool overwrite = true)
		{
			var list = operation.LastList;
			var resourceFolderPath = ResourcePath.FromDecodedUrl(list.RootFolder.Name + "/" + remotePath);

			var folder = list.RootFolder.Folders.AddUsingPath(resourceFolderPath, new FolderCollectionAddParameters { Overwrite = overwrite });

			folder.Context.Load(folder);

			return operation;
		}

		public static CSOMOperation DeleteFolder(this CSOMOperation operation, string remotePath)
		{
			var list = operation.LastList;
			var resourceFolderPath = ResourcePath.FromDecodedUrl(list.RootFolder.Name + "/" + remotePath);

			list.RootFolder.Folders.GetByPath(resourceFolderPath).DeleteObject();

			return operation;
		}
	}
}