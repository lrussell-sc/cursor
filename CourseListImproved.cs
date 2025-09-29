using System;
using System.Collections.Generic;
using System.Web.UI;
using Sitecore.Collections;
using Sitecore.Data.Items;
using ViridianSpark.Core;
using ViridianSpark.Core.Helpers;

namespace ViridianSpark.CourseCatalog
{
	public sealed class CourseListImproved : ViridianSpark.Core.WebControlBase
	{
		private const string FieldPrerequisites = "Prerequisites";
		private const string FieldCorequisites = "Corequisites";
		private const string FieldCrossListed = "Cross Listed Courses";
		private const string FieldCourseName = "Course Name";
		private const string FieldCourseDescription = "Course Description";
		private const string FieldCreditHoursNarrative = "Credit Hours Narrative";
		private const string NarrativeCoursesFolderName = "Narrative-Courses";

		protected override void DoRender(HtmlTextWriter output)
		{
			if (output == null)
			{
				return;
			}

			Item contextItem = Sitecore.Context.Item;
			if (contextItem == null)
			{
				return;
			}

			RenderCourses(contextItem, output);
		}

		private void RenderCourses(Item item, HtmlTextWriter output)
		{
			if (item == null)
			{
				return;
			}

			ChildList children = item.GetChildren();
			foreach (Item child in children)
			{
				if (child == null)
				{
					continue;
				}

				if (child.Parent != null && string.Equals(child.Parent.Name, NarrativeCoursesFolderName, StringComparison.Ordinal))
				{
					continue;
				}

				if (string.Equals(child.TemplateName, "Course", StringComparison.Ordinal))
				{
					output.AddAttribute(HtmlTextWriterAttribute.Class, "courseList");
					output.RenderBeginTag(HtmlTextWriterTag.Div);
					RenderAllCourseInformation(output, child);
					output.RenderEndTag(); // div
				}
				else
				{
					RenderCourses(child, output);
				}
			}
		}

		private void RenderAllCourseInformation(HtmlTextWriter output, Item course)
		{
			if (output == null || course == null)
			{
				return;
			}

			Dictionary<string, Item> prerequisites = GetDependenciesFromField(course, FieldPrerequisites);
			Dictionary<string, Item> corequisites = GetDependenciesFromField(course, FieldCorequisites);
			Dictionary<string, Item> crossListed = GetDependenciesFromField(course, FieldCrossListed);

			RenderCourseInformation(output, course);
			CourseCatalogUtilities.RenderCreditHours(output, course);
			RenderRequisites(output, course, prerequisites, corequisites, crossListed);
			RenderCredits(output, course);
		}

		private void RenderCourseInformation(HtmlTextWriter output, Item course)
		{
			output.RenderBeginTag(HtmlTextWriterTag.H2);

			string url = course.GetItemUrl();
			output.AddAttribute(HtmlTextWriterAttribute.Href, url);
			output.RenderBeginTag(HtmlTextWriterTag.A);

			output.RenderBeginTag(HtmlTextWriterTag.Span);
			string fullCourseNumber = course.GetFullCourseNumber(true) ?? string.Empty;
			output.WriteEncodedText(fullCourseNumber);
			output.RenderEndTag(); // span

			output.Write(" ");
			string courseName = course.GetFieldValue(FieldCourseName) ?? string.Empty;
			output.WriteEncodedText(courseName);

			output.RenderEndTag(); // a

			output.RenderEndTag(); // h2

			output.AddAttribute(HtmlTextWriterAttribute.Class, "desc");
			output.RenderBeginTag(HtmlTextWriterTag.Div);
			output.Write(CourseCatalogUtilities.FindCourseNameInTextAndReplaceWithLink(course, FieldCourseDescription));
			output.RenderEndTag(); // div
		}

		private void RenderCredits(HtmlTextWriter output, Item course)
		{
			string creditHours = course.GetFieldValue(FieldCreditHoursNarrative);
			if (!string.IsNullOrWhiteSpace(creditHours))
			{
				output.RenderBeginTag(HtmlTextWriterTag.H3);
				output.Write("Credits");
				output.RenderEndTag(); // h3
				output.Write(creditHours);
			}
		}

		private Dictionary<string, Item> GetDependenciesFromField(Item course, string listFieldName)
		{
			Dictionary<string, Item> requisites = new Dictionary<string, Item>(StringComparer.OrdinalIgnoreCase);
			List<Item> reqs = course.GetMultiListItem(listFieldName);

			if (reqs == null)
			{
				return requisites;
			}

			foreach (Item requisite in reqs)
			{
				if (requisite == null)
				{
					continue;
				}

				string courseNumber = requisite.GetFullCourseNumber(false);
				if (!string.IsNullOrWhiteSpace(courseNumber) && !requisites.ContainsKey(courseNumber))
				{
					requisites.Add(courseNumber, requisite);
				}
			}

			return requisites;
		}

		private void RenderRequisites(
			HtmlTextWriter output,
			Item course,
			Dictionary<string, Item> prerequisites,
			Dictionary<string, Item> corequisites,
			Dictionary<string, Item> crossListed)
		{
			RenderDependencies(output, course, "Prerequisite", prerequisites);
			RenderDependencies(output, course, "Corequisite", corequisites);
			RenderDependencies(output, course, "Cross Listed Courses", crossListed);
		}

		private void RenderDependencies(HtmlTextWriter output, Item course, string dependencyName, Dictionary<string, Item> requisites)
		{
			string narrativeFieldName = string.Concat(dependencyName, " Narrative");

			if (Sitecore.Context.PageMode.IsExperienceEditorEditing ||
				!string.IsNullOrEmpty(course.GetFieldValue(narrativeFieldName)))
			{
				RenderRequisiteHeading(output, dependencyName);
				RenderRequisiteNarrative(output, course, narrativeFieldName, requisites);
			}
		}

		private void RenderRequisiteHeading(HtmlTextWriter output, string dependencyName)
		{
			output.RenderBeginTag(HtmlTextWriterTag.H3);

			if (string.Equals(dependencyName, "Cross Listed Courses", StringComparison.Ordinal))
			{
				output.Write(dependencyName);
			}
			else
			{
				output.Write(string.Concat(dependencyName, "s"));
			}

			output.RenderEndTag(); // h3
		}

		private void RenderRequisiteNarrative(HtmlTextWriter output, Item course, string narrativeFieldName, Dictionary<string, Item> requisites)
		{
			if (Sitecore.Context.PageMode.IsExperienceEditorEditing)
			{
				output.Write(narrativeFieldName);
			}

			output.Write(CourseCatalogUtilities.FindCourseNameInTextAndReplaceWithLink(course, narrativeFieldName, requisites));
		}
	}
}

