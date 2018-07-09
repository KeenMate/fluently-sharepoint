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
		public static class DefaultRetrievals
		{
			public static Expression<Func<TermGroup, object>>[] TermGroup = new Expression<Func<TermGroup, object>>[]
			{
				g=>g.ContributorPrincipalNames,
				g=>g.Description,
				g=>g.GroupManagerPrincipalNames,
				g=>g.IsSiteCollectionGroup,
				g=>g.IsSystemGroup,
			};

			public static Expression<Func<TermSet, object>>[] TermSet = new Expression<Func<TermSet, object>>[]
			{
				s=>s.Contact,
				s=>s.Description,
				s=>s.IsOpenForTermCreation,
				s=>s.Names,
				s=>s.Stakeholders,
			};

			public static Expression<Func<Term, object>>[] Term = new Expression<Func<Term, object>>[]
			{
				t=>t.Description,
				t=>t.IsDeprecated,
				t=>t.IsKeyword,
				t=>t.IsRoot,
				t=>t.LocalCustomProperties,
				t=>t.CustomProperties,
				t=>t.TermsCount,
			};
		}
		

		private static void EnsureTaxonomyOperation(CSOMOperation operation)
		{
			operation.TaxonomyOperation = operation.TaxonomyOperation ?? new TaxonomyOperation(operation);
		}

		public static CSOMOperation LockTerm(this CSOMOperation operation)
		{
			operation.TaxonomyOperation.LevelLock = LockLevel.Term;
			return operation;
		}


		public static CSOMOperation OpenTaxonomySession(this CSOMOperation operation)
		{
			EnsureTaxonomyOperation(operation);

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
		public static CSOMOperation LoadTermGroup(this CSOMOperation operation, Guid guid, params Expression<Func<TermGroup, object>>[] retrievals)
		{
			var op = operation.TaxonomyOperation;

			op.LastTermGroup = op.LastTermStore.Groups.GetById(guid);
			operation.Load(op.LastTermGroup, retrievals.Length == 0 ? DefaultRetrievals.TermGroup : retrievals);
			
			return operation;
		}

		/// <summary>
		/// Load a term group from current <see cref="TaxonomyOperation.LastTermStore"/> by given <paramref name="name"/> parameter.
		/// </summary>
		/// <param name="operation"></param>
		/// <param name="name">Name of the term group</param>
		/// <param name="retrievals">When not retrievals are supplied <see cref="DefaultRetrievals.TermGroup"/> are used.</param>
		/// <returns></returns>
		public static CSOMOperation LoadTermGroup(this CSOMOperation operation, string name, params Expression<Func<TermGroup, object>>[] retrievals)
		{
			var op = operation.TaxonomyOperation;

			op.LastTermGroup = op.LastTermStore.Groups.GetByName(name);
			operation.Load(op.LastTermGroup, retrievals.Length == 0 ? DefaultRetrievals.TermGroup : retrievals);

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

		public static CSOMOperation LoadTermSet(this CSOMOperation operation, string name, params Expression<Func<TermSet, object>>[] retrievals)
		{
			var op = operation.TaxonomyOperation;
			var key = name.ToLower();
			op.LastTermSet = op.LastTermGroup.TermSets.GetByName(name);
			operation.Load(op.LastTermSet, retrievals.Length == 0 ? DefaultRetrievals.TermSet : retrievals);

			if (op.LoadedTermSets.ContainsKey(key))
			{
				op.LoadedTermSets[key] = op.LastTermSet;
			}
			else
			{
				op.LoadedTermSets.Add(key, op.LastTermSet);
			}

			return operation;
		}

		public static CSOMOperation LoadTerm(this CSOMOperation operation, string name, params Expression<Func<TermSet, object>>[] retrievals)
		{
			var op = operation.TaxonomyOperation;
			var key = name.ToLower();
			op.LastTerm = op.LastTermSet.Terms.GetByName(name);
			operation.Load(op.LastTermSet, retrievals.Length == 0 ? DefaultRetrievals.TermSet : retrievals);

			if (op.LoadedTermSets.ContainsKey(key))
			{
				op.LoadedTermSets[key] = op.LastTermSet;
			}
			else
			{
				op.LoadedTermSets.Add(key, op.LastTermSet);
			}

			return operation;
		}

		//public static CSOMOperation GetTermSets(this CSOMOperation operation, string name, params Expression<Func<TermSet, object>>[] retrievals)
		//{
		//	var op = operation.TaxonomyOperation;

		//	op.Session.GetTermSetsByName(new LabelMatchInformation(operation.Context)
		//	{
		//		StringMatchOption = StringMatchOption.StartsWith,
		//		TermLabel = name,

		//	});

		//	var key = name.ToLower();
		//	op.LastTermSet = op.LastTermGroup.TermSets.GetByName(name);
		//	operation.Load(op.LastTermSet, retrievals.Length == 0 ? DefaultRetrievals.TermSet : retrievals);

		//	if (op.LoadedTermSets.ContainsKey(key))
		//	{
		//		op.LoadedTermSets[key] = op.LastTermSet;
		//	}
		//	else
		//	{
		//		op.LoadedTermSets.Add(key, op.LastTermSet);
		//	}

		//	return operation;
		//}

	}
}