const btnO05 = document.getElementById("btn-o05");
const btnGG25 = document.getElementById("btn-gg25");
const resultsBox = document.getElementById("results-box");

btnO05.addEventListener("click", () => {
  resultsBox.innerHTML = `
    <div class="result-item">
      <div class="result-match">Modulo Over 0.5 PT pronto</div>
      <div class="result-badge o25">ATTIVO</div>
    </div>
  `;
});

btnGG25.addEventListener("click", () => {
  resultsBox.innerHTML = `
    <div class="result-item">
      <div class="result-match">Modulo GG / Over 2.5 pronto</div>
      <div class="result-badge gg">ATTIVO</div>
    </div>
  `;
});
