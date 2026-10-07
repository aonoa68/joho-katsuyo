window.MathJax = {
  tex: { inlineMath: [["\\(", "\\)"]], displayMath: [["\\[", "\\]"]], processEscapes: true },
  options: { ignoreHtmlClass: ".*", processHtmlClass: "arithmatex" }
};
document.addEventListener("DOMContentLoaded", function () {
  if (window.MathJax && MathJax.typesetPromise) MathJax.typesetPromise();
});
document$.subscribe(function () {
  if (window.MathJax && MathJax.typesetPromise) MathJax.typesetPromise();
});
