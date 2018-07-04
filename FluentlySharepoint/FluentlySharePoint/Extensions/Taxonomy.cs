using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using KeenMate.FluentlySharePoint.Models;
using KeenMate.FluentlySharePoint.Models.Taxonomy;
using Microsoft.SharePoint.Client.Taxonomy;

namespace KeenMate.FluentlySharePoint.Extensions
{
	public static class Taxonomy
	{
		private static void EnsureTaxonomyOperation(CSOMOperation operation)
		{
			operation.TaxonomyOperation = operation.TaxonomyOperation ?? new TaxonomyOperation();
		}

		public static CSOMOperation LockTerm(this CSOMOperation operation)
		{
			operation.TaxonomyOperation.LevelLock = LockLevel.Term;
			return operation;
		}


		public static CSOMOperation OpenTaxonomySession(this CSOMOperation operation)
		{
			if(operation.TaxonomyOperation == null)
				operation.TaxonomyOperation = new TaxonomyOperation();

			operation.TaxonomyOperation.Session = TaxonomySession.GetTaxonomySession(operation.Context);
			operation.Context.Load(operation.TaxonomyOperation.Session);

			return operation;
		}

		public static CSOMOperation SelectTaxonomyStore(this CSOMOperation operation, string storeName="", Guid? storeGuid = null)
		{
			var op = operation.TaxonomyOperation;
			if (!string.IsNullOrEmpty(storeName))
				op.LastTermStore = op.Session.TermStores.GetByName(storeName);
			else if (storeGuid.HasValue)
				op.LastTermStore = op.Session.TermStores.GetById(storeGuid.Value);
			else
				op.LastTermStore = op.Session.GetDefaultSiteCollectionTermStore();

			operation.Context.Load(op.LastTermStore);

			return operation;
		}

		public static CSOMOperation CreateTermGroup(this CSOMOperation operation, string name, Guid? guid = null)
		{
			var op = operation.TaxonomyOperation;

			op.LastTermGroup = op.LastTermStore.CreateGroup(name, guid ?? Guid.NewGuid());

			return operation;
		}

		public static CSOMOperation CreateTermSet(this CSOMOperation operation, string name, Guid? guid = null, int lcid = 1033)
		{
			var op = operation.TaxonomyOperation;

			op.LastTermSet = op.LastTermGroup.CreateTermSet(name, guid??Guid.NewGuid(), lcid);

			return operation;
		}

		public static CSOMOperation CreateTerm(this CSOMOperation operation, string name
			, Guid? guid = null, int lcid = 1033, IEnumerable<Tuple<string, int>> descriptions = null
			, Dictionary<string, string> customProperties = null, Dictionary<string, string> localProperties = null)
		{
			var op = operation.TaxonomyOperation;

			op.LastTerm = op.LastTermSet.CreateTerm(name, lcid, guid ?? Guid.NewGuid());

			if (descriptions != null)
				foreach (var description in descriptions)
				{
					op.LastTerm.SetDescription(description.Item1, description.Item2);
				}
			customProperties?.ToList().ForEach(x => op.LastTerm.SetCustomProperty(x.Key, x.Value));

			localProperties?.ToList().ForEach(x => op.LastTerm.SetLocalCustomProperty(x.Key, x.Value));

			return operation;
		}

		/// <summary>
		/// Load a term group from current <see cref="TaxonomyOperation.LastTermStore"/> by given <paramref name="guid"/> parameter.
		/// <para>This method does not add loaded term group in <see cref="TaxonomyOperation.LoadedTermGroups"/>.
		/// See Remarks section!</para>
		/// </summary>
		/// <param name="operation"></param>
		/// <param name="guid">Id of the term group</param>
		/// <remarks>This method does not add loaded group into <see cref="TaxonomyOperation.LoadedTermGroups"/> because the name is not available yet</remarks>
		/// <returns></returns>
		public static CSOMOperation LoadTermGroup(this CSOMOperation operation, Guid guid)
		{
			var op = operation.TaxonomyOperation;

			op.LastTermGroup = op.LastTermStore.Groups.GetById(guid);
			operation.Load(op.LastTermGroup);
			
			return operation;
		}

		/// <summary>
		/// Load a term group from current <see cref="TaxonomyOperation.LastTermStore"/> by given <paramref name="name"/> parameter.
		/// </summary>
		/// <param name="operation"></param>
		/// <param name="name">Name of the term group</param>
		/// <param name="retrivals"></param>
		/// <returns></returns>
		public static CSOMOperation LoadTermGroup(this CSOMOperation operation, string name, params Expression<Func<TermGroup, object>>[] retrivals)
		{
			var op = operation.TaxonomyOperation;

			op.LastTermGroup = op.LastTermStore.Groups.GetByName(name);
			operation.Load(op.LastTermGroup, retrivals);

			if (op.LoadedTermGroups.ContainsKey(name.ToLower()))
			{
				op.LoadedTermGroups[name.ToLower()] = op.LastTermGroup;
			}
			else
			{
				op.LoadedTermGroups.Add(name.ToLower(), op.LastTermGroup);
			}

			return operation;
		}

		public static CSOMOperation SelectTermGroup(this CSOMOperation operation, string name)
		{
			var op = operation.TaxonomyOperation;
			var key = name.ToLower();
			if (!op.LoadedTermGroups.ContainsKey(key))
			{
				throw new ArgumentOutOfRangeException(nameof(name), $"Could not find term group with name {key} in the already loaded groups.");
			}

			op.LastTermGroup = op.LoadedTermGroups[key];

			return operation;
		}

	}
}