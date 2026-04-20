import { describe, expect, it } from "vitest";
import { calendarDaysBetween, classifyMelTier, parseEticDate } from "../src/workOrderWatch";

describe("classifyMelTier", () => {
  it("detects below / at / above from phrases", () => {
    expect(classifyMelTier("Below MEL")).toBe("below");
    expect(classifyMelTier("at MEL")).toBe("at");
    expect(classifyMelTier("Above MEL")).toBe("above");
  });

  it("maps numeric MEL levels", () => {
    expect(classifyMelTier("1")).toBe("below");
    expect(classifyMelTier("3")).toBe("at");
    expect(classifyMelTier("4")).toBe("above");
  });

  it("returns unknown when empty", () => {
    expect(classifyMelTier("")).toBe("unknown");
  });
});

describe("parseEticDate", () => {
  it("parses ISO prefix and US dates", () => {
    expect(parseEticDate("2026-04-20")).toBe("2026-04-20");
    expect(parseEticDate("4/20/2026")).toBe("2026-04-20");
  });
});

describe("calendarDaysBetween", () => {
  it("counts whole days", () => {
    expect(calendarDaysBetween("2026-04-01", "2026-04-04")).toBe(3);
  });
});
