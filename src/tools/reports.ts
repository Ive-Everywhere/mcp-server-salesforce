import { Tool } from "@modelcontextprotocol/sdk/types.js";

export const MANAGE_REPORTS: Tool = {
  name: "salesforce_manage_reports",
  description: `Manage Salesforce Reports: list, describe metadata, execute (run), and retrieve async instance results.

Operations:
1. list - List recently viewed reports (up to 200) or query reports via SOQL for broader results
   - Use queryFilter to search reports by name, format, folder, or other Report object fields
   - Without queryFilter, returns recently viewed reports

2. describe - Get report metadata (columns, filters, groupings, report type) without executing
   - Requires reportId

3. execute - Run a report synchronously and get results
   - Requires reportId
   - Use includeDetails=true to get individual row data (not just aggregates)
   - Use filters to override report filters at runtime without modifying the saved report

4. executeAsync - Run a report asynchronously (for large/long-running reports)
   - Requires reportId
   - Returns an instanceId to retrieve results later

5. getInstances - List all async execution instances for a report
   - Requires reportId
   - Returns instance IDs with their status and completion timestamps

6. getInstanceResults - Retrieve results from an async report execution
   - Requires reportId and instanceId

Examples:
- List recent reports: { "operation": "list" }
- Search reports by name: { "operation": "list", "queryFilter": "Name LIKE '%Revenue%'" }
- Get report metadata: { "operation": "describe", "reportId": "00O..." }
- Run report with details: { "operation": "execute", "reportId": "00O...", "includeDetails": true }
- Run with filter override: { "operation": "execute", "reportId": "00O...", "filters": [{"column": "AMOUNT", "operator": "greaterThan", "value": "10000"}] }
- Async execution: { "operation": "executeAsync", "reportId": "00O...", "includeDetails": true }
- Get async results: { "operation": "getInstanceResults", "reportId": "00O...", "instanceId": "0LG..." }`,
  inputSchema: {
    type: "object",
    properties: {
      operation: {
        type: "string",
        enum: ["list", "describe", "execute", "executeAsync", "getInstances", "getInstanceResults"],
        description: "The report operation to perform"
      },
      reportId: {
        type: "string",
        description: "The Salesforce Report ID (required for describe, execute, executeAsync, getInstances, getInstanceResults)"
      },
      instanceId: {
        type: "string",
        description: "The async report instance ID (required for getInstanceResults)"
      },
      includeDetails: {
        type: "boolean",
        description: "Include detail rows in execution results (default: false, only aggregates returned)"
      },
      filters: {
        type: "array",
        items: {
          type: "object",
          properties: {
            column: { type: "string", description: "API name of the column to filter" },
            operator: { type: "string", description: "Filter operator: equals, notEqual, lessThan, greaterThan, lessOrEqual, greaterOrEqual, contains, notContain, startsWith, includes, excludes" },
            value: { type: "string", description: "Filter value" }
          },
          required: ["column", "operator", "value"]
        },
        description: "Runtime filter overrides for report execution (does not modify the saved report)"
      },
      queryFilter: {
        type: "string",
        description: "SOQL WHERE clause to filter reports in list operation (e.g., \"Name LIKE '%Revenue%'\" or \"Format = 'MATRIX'\")"
      },
      limit: {
        type: "number",
        description: "Maximum number of reports to return in list operation (default: 200)"
      }
    },
    required: ["operation"]
  }
};

export interface ReportFilter {
  column: string;
  operator: string;
  value: string;
}

export interface ManageReportsArgs {
  operation: 'list' | 'describe' | 'execute' | 'executeAsync' | 'getInstances' | 'getInstanceResults';
  reportId?: string;
  instanceId?: string;
  includeDetails?: boolean;
  filters?: ReportFilter[];
  queryFilter?: string;
  limit?: number;
}

/**
 * Format report fact map results into readable text
 */
function formatReportResults(result: any, includeDetails: boolean): string {
  const lines: string[] = [];

  // Report metadata summary
  const metadata = result.reportMetadata;
  if (metadata) {
    lines.push(`Report: ${metadata.name}`);
    lines.push(`Format: ${metadata.reportFormat}`);
    if (metadata.reportFilters?.length > 0) {
      lines.push(`Active Filters: ${metadata.reportFilters.length}`);
      for (const filter of metadata.reportFilters) {
        lines.push(`  - ${filter.column} ${filter.operator} ${filter.value}`);
      }
    }
    lines.push('');
  }

  // Extended metadata for column labels
  const extendedMeta = result.reportExtendedMetadata;
  const columnInfo = extendedMeta?.detailColumnInfo || {};
  const aggregateInfo = extendedMeta?.aggregateColumnInfo || {};

  // Fact map data
  const factMap = result.factMap;
  if (!factMap) {
    lines.push('No data returned.');
    return lines.join('\n');
  }

  // Grouping info
  const groupingsDown = result.groupingsDown?.groupings || [];
  const groupingsAcross = result.groupingsAcross?.groupings || [];

  if (groupingsDown.length > 0) {
    lines.push(`Row Groupings: ${groupingsDown.length}`);
  }
  if (groupingsAcross.length > 0) {
    lines.push(`Column Groupings: ${groupingsAcross.length}`);
  }

  // Grand totals (T!T key)
  if (factMap['T!T']) {
    const grandTotals = factMap['T!T'];
    if (grandTotals.aggregates?.length > 0) {
      lines.push('');
      lines.push('Grand Totals:');
      const aggKeys = Object.keys(aggregateInfo);
      for (let i = 0; i < grandTotals.aggregates.length; i++) {
        const agg = grandTotals.aggregates[i];
        const label = aggKeys[i] ? (aggregateInfo[aggKeys[i]]?.label || aggKeys[i]) : `Aggregate ${i}`;
        lines.push(`  ${label}: ${agg.label ?? agg.value ?? 'N/A'}`);
      }
    }
  }

  // Detail rows
  if (includeDetails && factMap['T!T']?.rows) {
    const rows = factMap['T!T'].rows;
    lines.push('');
    lines.push(`Detail Rows (${rows.length}):`);

    // Column headers
    const detailColumns = metadata?.detailColumns || [];
    const headerLabels = detailColumns.map((col: string) =>
      columnInfo[col]?.label || col
    );
    if (headerLabels.length > 0) {
      lines.push(`  | ${headerLabels.join(' | ')} |`);
      lines.push(`  | ${headerLabels.map(() => '---').join(' | ')} |`);
    }

    // Data rows
    for (const row of rows) {
      const cells = row.dataCells.map((cell: any) => cell.label ?? cell.value ?? '');
      lines.push(`  | ${cells.join(' | ')} |`);
    }
  }

  // Grouped results (for summary/matrix reports)
  if (groupingsDown.length > 0) {
    lines.push('');
    lines.push('Grouped Results:');
    formatGroupings(groupingsDown, factMap, aggregateInfo, columnInfo, metadata, includeDetails, lines, 0);
  }

  return lines.join('\n');
}

/**
 * Recursively format grouped report results
 */
function formatGroupings(
  groupings: any[],
  factMap: any,
  aggregateInfo: any,
  columnInfo: any,
  metadata: any,
  includeDetails: boolean,
  lines: string[],
  depth: number
): void {
  const indent = '  '.repeat(depth + 1);

  for (const group of groupings) {
    lines.push(`${indent}${group.label} (${group.value}):`);

    // Get subtotals for this group
    const key = `${group.key}!T`;
    const groupData = factMap[key];
    if (groupData?.aggregates?.length > 0) {
      const aggKeys = Object.keys(aggregateInfo);
      for (let i = 0; i < groupData.aggregates.length; i++) {
        const agg = groupData.aggregates[i];
        const label = aggKeys[i] ? (aggregateInfo[aggKeys[i]]?.label || aggKeys[i]) : `Aggregate ${i}`;
        lines.push(`${indent}  ${label}: ${agg.label ?? agg.value ?? 'N/A'}`);
      }
    }

    // Detail rows for this group
    if (includeDetails && groupData?.rows?.length > 0) {
      lines.push(`${indent}  Rows (${groupData.rows.length}):`);
      const detailColumns = metadata?.detailColumns || [];
      for (const row of groupData.rows) {
        const cells = row.dataCells.map((cell: any, idx: number) => {
          const colName = detailColumns[idx];
          const label = columnInfo[colName]?.label || colName || `Col${idx}`;
          return `${label}: ${cell.label ?? cell.value ?? ''}`;
        });
        lines.push(`${indent}    ${cells.join(', ')}`);
      }
    }

    // Recurse into sub-groupings
    if (group.groupings?.length > 0) {
      formatGroupings(group.groupings, factMap, aggregateInfo, columnInfo, metadata, includeDetails, lines, depth + 1);
    }
  }
}

export async function handleManageReports(conn: any, args: ManageReportsArgs) {
  const { operation, reportId, instanceId, includeDetails, filters, queryFilter, limit } = args;

  try {
    switch (operation) {
      case 'list': {
        if (queryFilter) {
          // Use SOQL to search reports with broader criteria
          let soql = `SELECT Id, Name, Description, Format, DeveloperName, FolderName, LastRunDate, CreatedDate, LastModifiedDate FROM Report WHERE ${queryFilter}`;
          if (limit) soql += ` LIMIT ${limit}`;
          else soql += ` LIMIT 200`;

          const result = await conn.query(soql);
          const reports = result.records.map((r: any, index: number) => {
            const parts = [
              `${index + 1}. ${r.Name}`,
              `   ID: ${r.Id}`,
              `   Developer Name: ${r.DeveloperName || 'N/A'}`,
              `   Format: ${r.Format || 'N/A'}`,
              `   Folder: ${r.FolderName || 'N/A'}`,
            ];
            if (r.Description) parts.push(`   Description: ${r.Description}`);
            if (r.LastRunDate) parts.push(`   Last Run: ${r.LastRunDate}`);
            return parts.join('\n');
          }).join('\n\n');

          return {
            content: [{
              type: "text",
              text: `Found ${result.records.length} reports:\n\n${reports}`
            }],
            isError: false,
          };
        } else {
          // Use Analytics API to get recently viewed reports
          const reports = await conn.analytics.reports();
          const formatted = reports.map((r: any, index: number) => {
            const parts = [
              `${index + 1}. ${r.name}`,
              `   ID: ${r.id}`,
            ];
            if (r.url) parts.push(`   URL: ${r.url}`);
            return parts.join('\n');
          }).join('\n\n');

          return {
            content: [{
              type: "text",
              text: `Recently viewed reports (${reports.length}):\n\n${formatted}`
            }],
            isError: false,
          };
        }
      }

      case 'describe': {
        if (!reportId) {
          return {
            content: [{ type: "text", text: "reportId is required for describe operation" }],
            isError: true,
          };
        }

        const report = conn.analytics.report(reportId);
        const meta = await report.describe();

        const lines: string[] = [];
        const rm = meta.reportMetadata;

        lines.push(`Report: ${rm.name}`);
        lines.push(`ID: ${rm.id}`);
        lines.push(`Format: ${rm.reportFormat}`);
        lines.push(`Report Type: ${rm.reportType?.type || 'N/A'}`);
        lines.push(`Description: ${rm.description || 'N/A'}`);

        // Detail columns
        if (rm.detailColumns?.length > 0) {
          lines.push('');
          lines.push('Detail Columns:');
          const detailInfo = meta.reportExtendedMetadata?.detailColumnInfo || {};
          for (const col of rm.detailColumns) {
            const info = detailInfo[col];
            if (info) {
              lines.push(`  - ${info.label} (${col}) [${info.dataType}]`);
            } else {
              lines.push(`  - ${col}`);
            }
          }
        }

        // Groupings
        if (rm.groupingsDown?.length > 0) {
          lines.push('');
          lines.push('Row Groupings:');
          for (const g of rm.groupingsDown) {
            lines.push(`  - ${g.name} (sort: ${g.sortOrder}, agg: ${g.dateGranularity || 'none'})`);
          }
        }
        if (rm.groupingsAcross?.length > 0) {
          lines.push('');
          lines.push('Column Groupings:');
          for (const g of rm.groupingsAcross) {
            lines.push(`  - ${g.name} (sort: ${g.sortOrder}, agg: ${g.dateGranularity || 'none'})`);
          }
        }

        // Filters
        if (rm.reportFilters?.length > 0) {
          lines.push('');
          lines.push('Filters:');
          for (const f of rm.reportFilters) {
            lines.push(`  - ${f.column} ${f.operator} ${f.value}`);
          }
        }

        // Aggregates
        const aggInfo = meta.reportExtendedMetadata?.aggregateColumnInfo;
        if (aggInfo && Object.keys(aggInfo).length > 0) {
          lines.push('');
          lines.push('Aggregates:');
          for (const [key, info] of Object.entries(aggInfo) as [string, any][]) {
            lines.push(`  - ${info.label} (${key}) [${info.dataType}]`);
          }
        }

        // Report type metadata
        const rtm = meta.reportTypeMetadata;
        if (rtm?.categories?.length > 0) {
          lines.push('');
          lines.push('Available Categories/Objects:');
          for (const cat of rtm.categories) {
            lines.push(`  - ${cat.label} (${cat.name})`);
          }
        }

        return {
          content: [{
            type: "text",
            text: lines.join('\n')
          }],
          isError: false,
        };
      }

      case 'execute': {
        if (!reportId) {
          return {
            content: [{ type: "text", text: "reportId is required for execute operation" }],
            isError: true,
          };
        }

        const report = conn.analytics.report(reportId);
        const options: any = {};

        if (includeDetails) {
          options.details = true;
        }

        // Apply runtime filter overrides
        if (filters && filters.length > 0) {
          options.metadata = {
            reportMetadata: {
              reportFilters: filters.map(f => ({
                column: f.column,
                operator: f.operator,
                value: f.value,
              }))
            }
          };
        }

        const result = await report.execute(options);
        const formatted = formatReportResults(result, includeDetails || false);

        return {
          content: [{
            type: "text",
            text: formatted
          }],
          isError: false,
        };
      }

      case 'executeAsync': {
        if (!reportId) {
          return {
            content: [{ type: "text", text: "reportId is required for executeAsync operation" }],
            isError: true,
          };
        }

        const report = conn.analytics.report(reportId);
        const options: any = {};

        if (includeDetails) {
          options.details = true;
        }

        if (filters && filters.length > 0) {
          options.metadata = {
            reportMetadata: {
              reportFilters: filters.map(f => ({
                column: f.column,
                operator: f.operator,
                value: f.value,
              }))
            }
          };
        }

        const instance = await report.executeAsync(options);

        return {
          content: [{
            type: "text",
            text: `Report execution started asynchronously.\n\nInstance ID: ${instance.id}\nStatus: ${instance.status}\nRequest Date: ${instance.requestDate || 'N/A'}\n\nUse operation "getInstanceResults" with this instanceId to retrieve results once complete.`
          }],
          isError: false,
        };
      }

      case 'getInstances': {
        if (!reportId) {
          return {
            content: [{ type: "text", text: "reportId is required for getInstances operation" }],
            isError: true,
          };
        }

        const report = conn.analytics.report(reportId);
        const instances = await report.instances();

        if (!instances || instances.length === 0) {
          return {
            content: [{
              type: "text",
              text: "No async report instances found for this report."
            }],
            isError: false,
          };
        }

        const formatted = instances.map((inst: any, index: number) => {
          return [
            `${index + 1}. Instance ID: ${inst.id}`,
            `   Status: ${inst.status}`,
            `   Request Date: ${inst.requestDate || 'N/A'}`,
            `   Completion Date: ${inst.completionDate || 'N/A'}`,
          ].join('\n');
        }).join('\n\n');

        return {
          content: [{
            type: "text",
            text: `Found ${instances.length} report instance(s):\n\n${formatted}`
          }],
          isError: false,
        };
      }

      case 'getInstanceResults': {
        if (!reportId) {
          return {
            content: [{ type: "text", text: "reportId is required for getInstanceResults operation" }],
            isError: true,
          };
        }
        if (!instanceId) {
          return {
            content: [{ type: "text", text: "instanceId is required for getInstanceResults operation" }],
            isError: true,
          };
        }

        const report = conn.analytics.report(reportId);
        const result = await report.instance(instanceId).retrieve();
        const formatted = formatReportResults(result, includeDetails || false);

        return {
          content: [{
            type: "text",
            text: formatted
          }],
          isError: false,
        };
      }

      default:
        return {
          content: [{
            type: "text",
            text: `Unknown report operation: ${operation}. Valid operations: list, describe, execute, executeAsync, getInstances, getInstanceResults`
          }],
          isError: true,
        };
    }
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    let enhancedError = errorMessage;

    if (errorMessage.includes('INSUFFICIENT_ACCESS') || errorMessage.includes('INVALID_CROSS_REFERENCE')) {
      enhancedError = `Access error: ${errorMessage}\n\nEnsure the user has:\n1. "Run Reports" permission\n2. Access to the report's folder\n3. The report ID is correct`;
    } else if (errorMessage.includes('INVALID_ID_FIELD') || errorMessage.includes('NOT_FOUND')) {
      enhancedError = `Report not found: ${errorMessage}\n\nCheck that the report ID is valid and starts with "00O".`;
    }

    return {
      content: [{
        type: "text",
        text: `Error in report ${operation}: ${enhancedError}`
      }],
      isError: true,
    };
  }
}
