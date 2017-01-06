class Filter {
  constructor (aqString) {
    this.reSingle = /(NOT )?@(\w+?)=="?([^\)]+)"?/;
    this.reMulti =  /(NOT )?@(\w+?)==\(([^\)]+)\)/;

    this.count = 0;
    this.map = {};
    this.original = '' + aqString;

    this.reduced = this.replaceFilterWithId(aqString);
  }
  create(f){
    let o = {};
    if (this.reMulti.test(f)) {
      if (RegExp.$1) {
        o.not = true;
      }
      o.type = 'multi';
      o.field = RegExp.$2;
      o.values = JSON.parse('[' + RegExp.$3 + ']');
    }
    else if (this.reSingle.test(f)) {
      if (RegExp.$1) {
        o.not = true;
      }
      o.type = 'single';
      o.field = RegExp.$2;
      let v = RegExp.$3;
      o.values = [ v.replace(/"$/g, '') ];
    }

    console.log(o);

    return o;
  }
  generateOutlookFilter() {
    let filters = this.reduced.match( /_F_\d+_/g );
    filters = (filters||[]).map(f=>{
      f = this.map[f];
      let field = this.getAlias(f.field);

      let values = f.values.map(v =>{
        return '(' + field + (f.not?' ne ':' eq ') + '\'' + v + '\')';
      });

      return (values.length > 1 ? '(' + values.join(' OR ') + ')' : values[0]);
    });
    return (filters.length > 1 ? '(' + filters.join(' AND ') + ')' : filters[0]);
  }
  getAlias(name) {
    return { From: 'From/EmailAddress/Address' }[name] || name;
  }
  getFields() {
    let o = {};
    for (var f in this.map) {
      o[ this.getAlias(this.map[f].field) ] = 1;
    }
    return Object.keys(o);
  }
  parse(re, aqString) {
    let m = re.exec(aqString);
    if (m) {
      let c = this.count++,
        id = '_F_'+c+'_';

      this.map[id] = this.create(m[0]);
      aqString = aqString.replace(m[0], id);
      aqString = this.reduce( aqString );
      return this.replaceFilterWithId(aqString);
    }
    return aqString;
  }
  reduce(f) {
    return f.replace( /\((_F_\d+_)\)/g, '$1' );
  }
  replaceFilterWithId(aqString) {
    aqString = this.parse(this.reMulti, aqString);
    aqString = this.parse(this.reSingle, aqString);

    return this.reduce(aqString);
  }
}

module.exports = Filter;
