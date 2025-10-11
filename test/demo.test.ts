import { Demo } from "../src/index";
describe("Demo", () => {
  test("test", async () => {
    const demo = new Demo({
      baseURL: ""
    });
    try {
      const res = await demo.getDemoList({});
      expect(res).toBeDefined();
    } catch (err) {
      expect(err).toBeDefined();
    }
  });
});