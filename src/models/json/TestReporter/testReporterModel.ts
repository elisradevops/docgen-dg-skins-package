export default class TestReporterModel {
  testReporter: any;

  constructor(data: any[], testPlanName: string) {
    this.testReporter = {
      type: 'testReporter',
      testPlanName: testPlanName,
      testSuites: data,
    };
  }

  getTestReporter(): any[] {
    return this.testReporter;
  }
}
