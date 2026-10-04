// Appraisal pricing, shared by portal.html (the Appraisal card) and
// client-form.html (the new-submission email). Answer keys match the name
// attributes in client-form.html.
(function () {
  const PRICE_PER_PAGE = 100;
  const PRICE_PER_FEATURE = 50;
  const PAGE_COUNT_KEY = 'Estimated Page Count';
  const PAGE_LIST_KEY = 'Anticipated Pages';
  const FEATURES_KEY = 'Interactive Features';
  const FEATURES_OTHER_KEY = 'Interactive Features (Other / Details)';

  const money = n => '$' + n.toLocaleString('en-US');

  // Uses the client's estimated page count; if they left it blank, counts
  // the pages they listed (split on commas, new lines, semicolons, bullets).
  function appraise(answers) {
    const a = answers || {};
    let pages = parseInt(a[PAGE_COUNT_KEY], 10);
    let pagesSource = 'estimated page count';
    if (!(pages > 0)) {
      pages = (a[PAGE_LIST_KEY] || '').split(/[,;\n•]+/).map(x => x.replace(/^[\s\-*\d.)]+/, '').trim()).filter(Boolean).length;
      pagesSource = pages ? 'counted from listed pages' : 'no pages given';
    }
    const features = (a[FEATURES_KEY] || '').split(', ').filter(f => f && f !== 'None needed');
    const pagesCost = pages * PRICE_PER_PAGE;
    const featuresCost = features.length * PRICE_PER_FEATURE;
    return {
      pages, pagesSource, features,
      other: (a[FEATURES_OTHER_KEY] || '').trim(),
      pricePerPage: PRICE_PER_PAGE, pricePerFeature: PRICE_PER_FEATURE,
      pagesCost, featuresCost, total: pagesCost + featuresCost
    };
  }

  window.Appraisal = { appraise, money, PRICE_PER_PAGE, PRICE_PER_FEATURE };
})();
