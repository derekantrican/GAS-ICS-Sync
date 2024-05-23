/*
Filters for calendar events based on ical properties (RFC 5545).
Define each filter with the following structure and add them to the var filters array:
{
  parameter: "property",      // Event property to filter by (e.g., "summary", "categories", "dtend", "dtstart").
  type: "include/exclude",    // Whether to include or exclude events matching the criteria.
  comparison: "method",       // Comparison method: "equals", "begins with", "contains", "regex", "<", ">".
                              // Note: "<", ">" only apply for date/time properties.
  criterias: ["values"],      // Array of values or patterns for comparison.
  offset: number              // (Optional) For date/time properties, specify an offset in days.
}
*/
var filters = [];

/* Examples:
var filters = [
  {
    parameter: "summary",       // Exclude events whose summary starts with "Pending:" or contains "cancelled".
    type: "exclude",
    comparison: "regex",
    criterias: ["^Pending:", "cancelled"]
  },
  {
    parameter: "categories",    // Include only events categorized as "Meetings".
    type: "include",
    comparison: "equals",
    criterias: ["Meetings"]
  },
  {
    parameter: "dtend",       // Reproduce the old onlyFutureEvents behaviour.
    type: "include",
    comparison: ">",
    offset: 0
  },
  {
    parameter: "dtstart",       // Exclude events starting more than 14 days from now.
    type: "exclude",
    comparison: ">",
    offset: 14
  }
];
*/