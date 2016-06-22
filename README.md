## Get related work items - VSTS Work Item Form Extension ##

A work item form extension that gets a list of related work items for the current work item opened in work item form.

#### Overview ####

This extension adds a page to work item form which shows a list of related work items based on a list of configurable fields. The extension retrieves the field values from the current work item in the form and load a list of related work items based on those values. Users can configure which fields do they want to look up for to retrieve related work items and which field to use to sort the work items list. By default, the extension loads top 20 workitems based on the search criteria.

![Group](img/Example.png)

In the screenshot above, the extension loads work items which share same field values as "Work Item type", "State", "Area Path" and "Tags", sorted by ChangedDate field. It does exactly what a VSTS query does, but in a much simpler and faster way.
Users can choose to change the look up fields and "order by" field and save these settings for the current work item type. The next time they open up the form, the extension will read user's settings and retrieve the work item list based on them.

![Group](img/AddLinkExample.png)

Users can also add a link to any of the workitems in the list directly from the extension by clicking on the link icon on the left of work item type color bar, as shown in the screenshot above.