function formatShortDate(value) {
    const date =
        value instanceof Date
            ? value
            : dateValue(value);

    if (!date) return "—";

    return date.toLocaleDateString(
        "en-IN",
        {
            day: "2-digit",
            month: "short"
        }
    );
}
