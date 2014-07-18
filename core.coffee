class Core
  debug: false

  p: (msg) -># {{{
    WScript.Echo msg# }}}

  dp: (msg) -># {{{
    WScript.Echo msg if @debug# }}}

  filterArray: (array, filter) -># {{{
    @dp "filter = [#{filter}]"
    ret = []
    for a in array
      if a.match eval "/#{filter}/"
        @dp a
        ret.push a
    return ret# }}}
