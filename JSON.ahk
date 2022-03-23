/*
Thanks Bentschi
https://www.autohotkey.com/board/topic/91700-funktion-json-encode-und-decode/
*/

json(i)
{
  ;ENCODE
  if (isobject(i))
  {
    o := "", a := 1, x := 1
    for k,v in i
    {
      if (k!=x)
        a := 0, break
      x += 1
    }
    o .= (a) ? "[" : "{", f := 1
    for k,v in i
      o .= ((f) ? "" : ",")((a) ? "" : """" k """:")((isobject(v)) ? json(v) : ((v+0=v) ? v : """" v """")), f := 0
    return o ((a) ? "]" : "}")
  }
  ;DECODE
  if (regexmatch(i, "s)^__chr(A|W):(.*)", m))
  {
    VarSetCapacity(b, 4, 0), NumPut(m2, b, 0, "int")
    return StrGet(&b, 1, (m1="A") ? "cp28591" : "utf-16")
  }
  if (regexmatch(i, "s)^__str:((\\""|[^""])*)", m))
  {
    str := m1
    for p,r in {b:"`b", f:"`f", n:"`n", 0:"", r:"`r", t:"`t", v:"`v", "'":"'", """":"""", "/":"/"}
      str := regexreplace(str, "\\" p, r)
    while (regexmatch(str, "s)^(.*?)\\x([0-9a-fA-F]{2})(.*)", m))
      str := m1 json("__chrA:0x" m2) m3
    while (regexmatch(str, "s)^(.*?)\\u([0-9a-fA-F]{4})(.*)", m))
      str := m1 json("__chrW:0x" m2) m3
    while (regexmatch(str, "s)^(.*?)\\([0-9]{1,3})(.*)", m))
      str := m1 json("__chrA:" m2) m3
    return regexreplace(str, "\\\\", "\")
  }
  str := [], obj := []
  while (RegExMatch(i, "s)^(.*?[^\\])""((\\""|[^""])*?[^\\]|)""(.*)$", m))
    str.insert(json("__str:" m2)), i := m1 "__str<" str.maxIndex() ">" m4
  while (RegExMatch(RegExReplace(i, "\s+", ""), "s)^(.*?)(\{|\[)([^\{\[\]\}]*?)(\}|\])(.*)$", m))
  {
    a := (m2="{") ? 0 : 1, c := m3, i := m1 "__obj<" ((obj.maxIndex()+1) ? obj.maxIndex()+1 : 1) ">" m5, tmp := []
    while (RegExMatch(c, "^(.*?),(.*)$", m))
      tmp.insert(m1), c := m2
    tmp.insert(c), tmp2 := {}, obj.insert(cobj := {})
    for k,v in tmp
    {
      if (RegExMatch(v, "^(.*?):(.*)$", m))
        tmp2[m1] := m2
      else
        tmp2.insert(v)
    }
    for k,v in tmp2
    {
      for x,y in str
        k := RegExReplace(k, "__str<" x ">", y), v := RegExReplace(v, "__str<" x ">", y)
      for x,y in obj
        v := RegExMatch(v, "^__obj<" x ">$") ? y : v
      cobj[k] := v
    }
  }
  return obj[obj.maxIndex()]
}