let filters = [
  '@From=="erocheleau@coveo.com"',
  '(NOT @From=="bphillips@coveo.com")',
  '@From==("erocheleau@coveo.com","bphillips@coveo.com")',
  '(@From=="erocheleau@coveo.com") ((NOT @From=="bphillips@coveo.com"))',
  '(@From==("kbarnhart@coveo.com","mpumper@coveo.com")) ((NOT @From==("erocheleau@coveo.com","bphillips@coveo.com")))',
  '((NOT @From==("erocheleau@coveo.com","bphillips@coveo.com"))) (@From==("kbarnhart@coveo.com","mpumper@coveo.com"))',
  '@Importance==High',
  '(@From=="erocheleau@coveo.com") (@Importance==Low)'
];

let Filter = require('./src/Filter');

filters.forEach(f=>{
  console.log('\n\n------');
  let f1 = new Filter(f);
  console.log(f1.generateOutlookFilter() );
  console.log( f1.getFields() );
});
