// Utility functions

// Finds difference between two dates in hours (2 decimal places)
export function DateDiffHrs(from: Date, to: Date) : number {

    return parseFloat(((from.getTime() - to.getTime())/3600000).toFixed(2));
}