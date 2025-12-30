/**
 * Ribbon button command handlers
 */
/**
 * Show taskpane when ribbon button is clicked
 */
declare function showTaskpane(event: Office.AddinCommands.Event): void;
/**
 * Quick action: Check classification status
 */
declare function checkStatus(event: Office.AddinCommands.Event): Promise<void>;
/**
 * Classify current email manually
 */
declare function classifyEmail(event: Office.AddinCommands.Event): Promise<void>;
//# sourceMappingURL=commands.d.ts.map