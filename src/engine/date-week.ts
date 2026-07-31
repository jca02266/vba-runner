/** Shared VBA DatePart/Format week-number calculation. */
export function vbaWeekNumber(date: Date, firstDayOfWeek: number, firstWeekOfYear: number): number {
    const weekStart = firstDayOfWeek <= 1 ? 0 : firstDayOfWeek - 1;
    const firstWeekStart = (year: number): Date => {
        const jan1 = new Date(year, 0, 1);
        const offset = (jan1.getDay() - weekStart + 7) % 7;
        const start = new Date(year, 0, 1 - offset);
        if (firstWeekOfYear === 1) return start;
        const daysInNewYear = 7 - offset;
        const requiredDays = firstWeekOfYear === 2 ? 4 : 7;
        if (daysInNewYear >= requiredDays) return start;
        start.setDate(start.getDate() + 7);
        return start;
    };
    let start = firstWeekStart(date.getFullYear());
    if (date < start) start = firstWeekStart(date.getFullYear() - 1);
    const dayStart = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const startDay = new Date(start.getFullYear(), start.getMonth(), start.getDate());
    return Math.floor((Date.UTC(dayStart.getFullYear(), dayStart.getMonth(), dayStart.getDate()) -
        Date.UTC(startDay.getFullYear(), startDay.getMonth(), startDay.getDate())) / 86400000 / 7) + 1;
}
