describe("empty spec", () => {
  it("passes", () => {
    cy.visit("http://localhost:3000/spreadsheet");
    cy.wait(5000); // 値が設定されるのを待つ APIを呼んで値を設定する場合はhttpリクエストをwaitでもOK
    cy.get(".igr-spreadsheet").should("be.visible");
    cy.contains("I1234");
    cy.contains("メーカー名");
  });
});
