import { getAllSubjects } from "@/lib/course-utils";
import { describe, expect, it } from "vitest";

describe("course-utils", () => {
  describe("getAllSubjects", () => {
    it("Should return and array", () => {
      const result = getAllSubjects();
      expect(Array.isArray(result)).toBe(true);
    });
  });
});
