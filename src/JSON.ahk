/**
 * @author GroggyOtter <groggyotter@gmail.com>
 * @version 1.1
 * @see https://github.com/GroggyOtter/jsongo_AHKv2
 * @license GNU
 * @classdesc Library for conversion of JSON text to AHK object and vice versa
 * 
 * @property {number} escape_slash     - If true, adds the optional escape character to forward slashes
 * @property {number} escape_backslash - If true, backslash is encoded as `\\` otherwise it is encoded as `\u005C`
 * @property {number} inline_arrays    - If true, arrays containing only strings/numbers are kept on 1 line
 * @property {number} extract_objects  - If true, attempts to extract literal objects instead of erroring
 * @property {number} extract_all      - If true, attempts to extract all object types instead of erroring
 * @property {number} silent_error     - If true, error popups are supressed and are instead written to the .error_log property
 * @property {number} error_log        - Stores error messages when an error occurs and the .silent_error property is true
 */
class jsongo {
    #Requires AutoHotkey 2.0.11+
    static version := '1.1'
    
    ; === User Options ===
    /**
     * @description If true, adds the optional escape character to forward slashes  
     * @type {(Boolean)}
     * @example
     * ; Using the string http://autohotkey.com
     * jsongo.escape_slash := 0 ; => http://autohotkey.com
     * jsongo.escape_slash := 1 ; => http:\/\/autohotkey.com
     */
    static escape_slash := 1
    
    /**
     * @description If true, backslash is encoded as `\\` otherwise it is encoded as `\u005C`
     * @type {(Boolean)}
     * @example
     * ; C:\Some\Path\file.exe
     * jsongo.escape_backslash := 0 ; => C:\u005CSome\u005CPath\u005Cfile.exe
     * jsongo.escape_backslash := 1 ; => C:\\Some\\Path\\file.exe
     */
    static escape_backslash := 1
    
    /**
     * @description If true, arrays containing only strings/numbers are kept on 1 line
     * @type {(Boolean)}
     * @example
     * ; C:\Some\Path\file.exe
     * jsongo.inline_arrays := 0 ; => my_array: [
     * ;                                  1,
     * ;                                  2,
     * ;                                  3
     * ;                              ]
     * jsongo.inline_arrays := 1 ; => my_array: [1, 2, 3]
     */
    static inline_arrays := 0
    
    /**
     * @description If true, attempts to extract literal objects instead of erroring
     * @type {(Boolean)}
     */
    static extract_objects := 1
    
    /**
     * @description If true, attempts to extract all object types instead of erroring
     * @type {(Boolean)}
     */
    static extract_all := 1
    
    /**
     * @description If true, error popups are supressed and are instead written to the .error_log property
     * @type {(Boolean)}
     */
    static silent_error := 1
    
    /**
     * @description Stores error messages when an error occurs and the .silent_error property is true
     * @type {(String)}
     */
    static error_log := ''
    
    ; === User Methods ===
    /**
     * @description Converts a string of JSON text into an AHK object
     * @method
     * @param {(String)} jtxt JSON string to convert into an AHK [object](https://www.autohotkey.com/docs/v2/lib/Object.htm)  
     * @param {(FuncObj)} [reviver=''] [optional] Reference to a reviver function.  
     * A reviver function receives each key:value pair before being added to the object and must have at least 3 parameters.  
     * @returns {(Map|Array|String)} Return type is based on JSON text input.  
     * It will be map, array, or string.  
     * On failure, an error message is thrown or an empty string is returned if `.silent_error` is true.
     * @access public
     * @example <caption>  
     * Converting JOSN text to object.</caption>
     * txt := '{"a":1, "b":2}'
     * obj := jsongo.Parse(txt)
     * MsgBox(obj['b']) ; Shows 2
     */
    static Parse(jtxt, reviver:='') => this._Parse(jtxt, reviver)
    
    /**
     * @description Converts a string of JSON text into an AHK object
     * @method
     * @param {(Map|Array)} base_item - A map or array to convert into JSON format.  
     * If the `.extract_objects` property is true, literal objects are also accepted.  
     * If the `.extract_all` property or the `extract_all` parameter are true, all object types are accepted.  
     * @param {(FuncObj)} [replacer=''] - Reference to a replacer function.  
     * A replacer function receives each key:value pair before being added to the JSON string.  
     * The function must have at least 3 parameters to receive the key, the value, and the removal variable.  
     * @param {(String|Number)} [spacer=''] - Defines the character set used to indent each level of the JSON tree.  
     * Number indicates the number of spaces to use for each indent.  
     * String indiciates the characters to use. `` `t `` would be 1 tab for each indent level.  
     * If omitted or an empty string is passed in, the JSON string will export as a single line of text.  
     * @param {(Number)} [extract_all=0] - If true, `base_item` can be any object type instead of throwing an error.
     * @returns {(String)} Returns a JSON string.  
     * On failure, an error message is thrown or an empty string is returned if `.silent_error` is true.
     * @access public
     * @example <caption>  
     * Converting object JSON text.</caption>
     * obj := Map('a', [1,2,3], 'b', [4,5,6])
     * json := jsongo.Stringify(obj, , 4)
     * MsgBox(json)
     */
    static Stringify(base_item, replacer:='', spacer:='', extract_all:=0) => this._Stringify(base_item, replacer, spacer, extract_all)
    
    /**
     * @access private
     */
    static _Parse(jtxt, reviver:='') {
        static xval := 1, xobj := 2, xarr := 3, xkey := 4, xstr := 5, xend := 6, xcln := 7, xeof := 8
        , xerr := 9, null := '', str_flag := Chr(5), tmp_q := Chr(6), tmp_bs:= Chr(7), fn := A_ThisFunc
        
        expect   := xval
        , json   := []
        , path   := [json]
        , key    := ''
        , is_key := 0
        , remove := jsongo.JSON_Remove()
        this.error_log := ''
        if_rev := (reviver is Func && reviver.MaxParams > 2) ? 1 : 0
        
        ; Illegal control char check
        loop 31
            switch A_Index {
                case 9, 10, 13: continue
                default:(i := InStr(jtxt, Chr(A_Index), 1))
                    ? err(21, i, 'Character number: 9, 10, 13 or anything higher than 31.', A_Index)
                    : 0
            }
        
        ; Escape chars 
        esc_arr :=
                [['\u005C', tmp_bs]
                ,['\\'    , tmp_bs]
                ,['\"'    ,tmp_q]
                ,['"'     ,str_flag]
                ,[tmp_q   ,'"']
                ,['\/'    ,'/']
                ,['\b'    ,'`b']
                ,['\f'    ,'`f']
                ,['\n'    ,'`n']
                ,['\r'    ,'`r']
                ,['\t'    ,'`t'] ]
        
        for k, esc in esc_arr
            this.replace_if_exist(&jtxt, esc[1], esc[2])
        
        ; Unicode
        i := 0
        while (i := InStr(jtxt, '\u', 1, ++i))
            IsNumber('0x' (hex := SubStr(jtxt, i+2, 4)))
                ? jtxt := StrReplace(jtxt, '\u' hex, Chr(('0x' hex)), 1)
                : err(22, i+2, '\u0000 to \uFFFF', '\u' hex)
        
        ; Illegal escape char check
        if (i := InStr(jtxt, '\', 1))
            err(23, i+1, '\b \f \n \r \t \" \\ \/ \u', '\' SubStr(jtxt, i+1, 1))
        else jtxt := StrReplace(jtxt, tmp_bs, '\', 1)
        
        ; Get length and set index to 1
        jlength := StrLen(jtxt) + 1
        , ji := 1
        
        ; Parse text
        while (ji < jlength) {
            ; Skip valid whitespace
            if InStr(' `t`n`r', (char := SubStr(jtxt, ji, 1)), 1)
                ji++
            ; Process text
            else switch expect {
                ; JSON value
                case xval:
                    valuelabel:
                    (char == '{')
                        ? (o := Map(), (path[path.Length] is Array)
                            ? path[path.Length].Push(o)
                        : path[path.Length][key] := o, path.Push(o), expect := xobj, ji++)
                    : (char == '[')
                        ? (a := [], (path[path.Length] is Array)
                            ? path[path.Length].Push(a)
                        : path[path.Length][key] := a, path.Push(a), expect := xarr, ji++)
                    : (char == str_flag)
                        ? (end := InStr(jtxt, str_flag, 1, ji+1))
                            ? is_key
                                ? (is_key := 0, key := SubStr(jtxt, ji+1, end-ji-1), expect := xcln, ji := end+1)
                            : (rev(SubStr(jtxt, ji+1, end-ji-1)), expect := xend, ji := end+1)
                        : err(24, ji, '"', SubStr(jtxt, ji))
                    : InStr('-0123456789', char, 1)
                       ? RegExMatch(jtxt, '(-?(?:0|[123456789]\d*)(?:\.\d+)?(?:[eE][-+]?\d+)?)', &match, ji)
                            ? (rev(Number(match[])), expect := xend, ji := match.Pos + match.Len )
                        : err(25, ji, , SubStr(jtxt, ji))
                    : (char == 't')
                        ? (SubStr(jtxt, ji, 4) == 'true')
                            ? (rev(true) , ji += 4, expect := xend)
                        : err(26, ji + tfn_idx('true', SubStr(jtxt, ji, 4)), 'true' , SubStr(jtxt, ji, 4))
                    : (char == 'f')
                        ? (SubStr(jtxt, ji, 5) == 'false')
                            ? (rev(false), ji += 5, expect := xend)
                        : err(27, ji + tfn_idx('false', SubStr(jtxt, ji, 5)), 'false', SubStr(jtxt, ji, 5))
                    : (char == 'n')
                        ? (SubStr(jtxt, ji, 4) == 'null')
                            ? (rev(null) , ji += 4, expect := xend)
                        : err(28, ji + tfn_idx('null', SubStr(jtxt, ji, 4)), 'null' , SubStr(jtxt, ji, 4))
                    : err(29, ji, '`n`tArray: [ `n`tObject: { `n`tString: " `n`tNumber: -0123456789 `n`ttrue/false/null: tfn ', char)
                ; JSON array
                case xarr:
                    if (char == ']')
                        path_pop(&char), expect := (path.Length = 1)
                            ? xeof
                        : xend, ji++
                    else goto('valuelabel')
                ; JSON object
                case xobj: 
                    switch char {
                        case str_flag: goto((is_key := 1) ? 'valuelabel' : 0)
                        case '}': path_pop(&char), expect := (path.Length = 1) ? xeof : xend, ji++
                        default: err(31, ji, '"}', char)
                    }
                ; JSON object key
                case xkey:
                    if (char == str_flag)
                        goto((is_key := 1) ? 'valuelabel' : 0)
                    else err(32, ji, '"', char)
                ; JSON object separator
                case xcln: (char == ':') ? (expect := xval, ji++) : err(33, ji, ':', char)
                ; JSON end of element
                case xend: (char == ',') ? (ji++, expect := (path[path.Length] is Array) ? xval : xkey)
                    ; End of object
                    : (char == '}')
                        ? (ji++, (path[path.Length] is Map)
                            ? path_pop(&char)
                        : err(34, ji, ']', char), (path.Length = 1) ? expect := xeof : 0`)
                    ; End of array
                    : (char == ']')
                        ? (ji++, (path[path.Length] is Array)
                            ? path_pop(&char)
                        : err(35, ji, '}', char), (path.Length = 1) ? expect := xeof : 0`)
                    : err(36, ji, '`nEnd of array: ]`nEnd of object: }`nNext value: ,`nWhitespace: [Space] [Tab] [Linefeed] [Carriage Return]', char)
                ; JSON 
                case xeof: err(40, ji, 'End of JSON', char)
                ; JSON error
                case xerr: return ''
            }
        }
        
        ; Ensure all objects and arrays were successfully closed and return the converted JSON data
        return (path.Length != 1) ? err(37, ji, 'Size: 1', 'Actual size: ' path.Length) : json[1]
        
        ; Remove the current path
        path_pop(&char) => (path.Length > 1) ? path.Pop() : err(38, ji, 'Size > 0', 'Actual size: ' path.Length-1)
        ; Reviver
        rev(value) => (path[path.Length] is Array) ? (if_rev ? value := reviver((path[path.Length].Length), value, remove) : 0, (value == remove) ? '' : path[path.Length].Push(value) ) : (if_rev ? value := reviver(key, value, remove) : 0, (value == remove) ? '' : path[path.Length][key] := value )
        ; Error handling
        err(msg_num, idx, ex:='', rcv:='') {
            clip := '`n'
            ,offset := 50
            ,clip := 'Error Location:`n', clip .= (idx > 1) ? SubStr(jtxt, 1, idx-1) : ''
            ,(StrLen(clip) > offset)
                ? clip := SubStr(clip, (offset * -1))
            : 0,  clip .= '>>>' SubStr(jtxt, idx, 1) '<<<',  post_clip := (idx < StrLen(jtxt))
                ? SubStr(jtxt, idx+1)
            : '',  clip .= (StrLen(post_clip) > offset)
                ? SubStr(post_clip, 1, offset)
            : post_clip,  clip := StrReplace(clip, str_flag, '"')
            ,this.error(msg_num, fn, ex, rcv, clip), expect := xerr
        }
        ; Error handling for true/false/null
        tfn_idx(a, b) {
            loop StrLen(a)
                if SubStr(a, A_Index, 1) !== SubStr(b, A_Index, 1)
                    Return A_Index-1
        }
    }
    
    /**
     * @access private
     */
    static _Stringify(base_item, replacer, spacer, extract_all) {
        static fn := A_ThisFunc
        
        ; Replacer validation
        switch Type(replacer) {
            case 'Func': if_rep := (replacer.MaxParams > 2) ? 1 : 0
            case 'Array':
                if_rep := 2, omit := Map(), omit.Default := 0
                for i, v in replacer
                    omit[v] := 1
            default: if_rep := 0
        }
        
        ; Spacer validation
        switch Type(spacer) {
            case 'String': _ind := spacer, lf := (spacer == '') ? '' : '`n'
                if (spacer == '')
                    _ind := lf := '', cln := ':'
                else _ind := spacer, lf := '`n', cln := ': '
            case 'Integer','Float','Number':
                lf := '`n', cln := ': ', _ind := ''
                loop Floor(spacer)
                    _ind .= ' '
            default: _ind := lf := '', cln := ':'
        }
        
        this.error_log := ''
        , extract_all := (extract_all) ? 1 
            : this.extract_all ? 1
            : 0
        , remove := jsongo.JSON_Remove()
        , value_types := 'String Number Array Map'
        , value_types .= extract_all
            ? ' AnyObject'
            : this.extract_objects
                ? ' LiteralObject' 
                : ''
        
        (if_rep = 1) ? base_item := replacer('', base_item, remove) : 0
        if (base_item = remove)
            return ''
        else jtxt := extract_data(base_item)
        
        ; Encode control characters and slashes
        loop 33
            switch A_Index {
                case 9,10,13: continue
                case  8: this.replace_if_exist(&jtxt, Chr(A_Index), '\b')
                case 12: this.replace_if_exist(&jtxt, Chr(A_Index), '\f')
                case 32: (this.escape_slash) ? this.replace_if_exist(&jtxt, '/', '\/') : 0
                case 33: (this.escape_backslash) ? this.replace_if_exist(&jtxt, '\u005C', '\\') : 0 
                default: this.replace_if_exist(&jtxt, Chr(A_Index), Format('\u{:04X}', A_Index))
            }
        
        return jtxt
        
        ; Function that recursively extracts data from objects
        extract_data(item, ind:='') {
            switch Type(item) {
                case 'String': return '"' encode(&item) '"'
                case 'Integer','Float': return item
                case 'Array':
                    str := '['
                    if (ila := this.inline_arrays ?  1 : 0)
                        for i, v in item
                            InStr('String|Float|Integer', Type(v), 1) ? 1 : ila := ''
                        until (!ila)
                    for i, v in item
                        (if_rep = 2 && omit[i]) ? '' : (if_rep = 1 && (v := replacer(i, v, remove)) = remove) ? '' : str .= (ila ? extract_data(v, ind _ind) ', ' : lf ind _ind extract_data(v, ind _ind) ',')
                    return ((str := RTrim(str, ', ')) == '[') ? '[]' : str (ila ? '' : lf ind) ']'
                case 'Map':
                    str := '{'
                    for k, v in item
                        (if_rep = 2 && omit[k]) ? '' : (if_rep = 1 && (v := replacer(k, v, remove)) = remove) ? '' : str .= lf ind _ind (k is String ? '"' encode(&k) '"' cln : err(11, 'String', Type(k))) extract_data(v, ind _ind) ','
                    return ((str := RTrim(str, ',')) == '{') ? '{}' : str lf ind '}'
                case 'Object':
                    (this.extract_objects) ? 1 : err(12, value_types, Type(item))
                    Object:
                    str := '{'
                    for k, v in item.OwnProps()
                        (if_rep = 2 && omit[k]) ? '' : (if_rep = 1 && (v := replacer(k, v, remove)) = remove) ? '' : str .= lf ind _ind (k is String ? '"' encode(&k) '"' cln : err(11, 'String', Type(k))) extract_data(v, ind _ind) ','
                    return ((str := RTrim(str, ',')) == '{') ? '{}' : str lf ind '}'
                case 'VarRef','ComValue','ComObjArray','ComObject','ComValueRef': return err(15, 'These are not of type "Object":`nVarRef ComValue ComObjArray ComObject and ComValueRef', Type(item))
                default:
                    !extract_all ? err(13, value_types, Type(item)) : 0
                    goto('Object')
            }
        }
        
        encode(&str) => (this.replace_if_exist(&str ,  '\', '\u005C'), this.replace_if_exist(&str,  '"', '\"'), this.replace_if_exist(&str, '`t', '\t'), this.replace_if_exist(&str, '`n', '\n'), this.replace_if_exist(&str, '`r', '\r')) ? str : str
        err(msg_num, ex:='', rcv:='') => this.error(msg_num, fn, ex, rcv)
    }

    /**
     * @access private
     */
    class JSON_Remove {
    }
    
    /**
     * @access private
     */
    static replace_if_exist(&txt, find, replace) => (InStr(txt, find, 1) ? txt := StrReplace(txt, find, replace, 1) : 0)
    
    /**
     * @access private
     */
    static error(msg_num, fn, ex:='', rcv:='', extra:='') {
        err_map := Map(11,'Stringify error: Object keys must be strings.'  ,12,'Stringify error: Literal objects are not extracted unless:`n-The extract_objects property is set to true`n-The extract_all property is set to true`n-The extract_all parameter is set to true.'  ,13,'Stringify error: Invalid object found.`nTo extract all objects:`n-Set the extract_all property to true`n-Set the extract_all parameter to true.'  ,14,'Stringify error: Invalid value was returned from Replacer() function.`nReplacer functions should always return a string or the "remove" value passed into the 3rd parameter.'  ,15,'Stringify error: Invalid object encountered.'  ,21,'Parse error: Forbidden character found.`nThe first 32 ASCII chars are forbidden in JSON text`nTab, linefeed, and carriage return may appear as whitespace.'  ,22,'Parse error: Invalid hex found in unicode escape.`nUnicode escapes must be in the format \u#### where #### is a hex value between 0000 and FFFF.`nHex values are not case sensitive.'  ,23,'Parse error: Invalid escape character found.'  ,24,'Parse error: Could not find end of string'  ,25,'Parse error: Invalid number found.'  ,26,'Parse error: Invalid `'true`' value.'  ,27,'Parse error: Invalid `'false`' value.'  ,28,'Parse error: Invalid `'null`' value.'  ,29,'Parse error: Invalid value encountered.'  ,31,'Parse error: Invalid object item.'  ,32,'Parse error: Invalid object key.`nObject values must have a string for a key name.'  ,33,'Parse error: Invalid key:value separator.`nAll keys must be separated from their values with a colon.'  ,34,'Parse error: Invalid end of array.'  ,35,'Parse error: Invalid end of object.'  ,36,'Parse error: Invalid end of value.'  ,37,'Parse error: JSON has objects/arrays that have not been terminated.'  ,38,'Parse error: Cannot remove an object/array that does not exist.`nThis error is usually thrown when there are extra closing brackets (array)/curly braces (object) in the JSON string.'  ,39,'Parse error: Invalid whitespace character found in string.`nTabs, linefeeds, and carriage returns must be escaped as \t \n \r (respectively).'  ,40,'Characters appears after JSON has ended.' )
        msg := err_map[msg_num], (ex != '') ? msg .= '`nEXPECTED: ' ex : 0, (rcv != '') ? msg .= '`nRECEIVED: ' rcv : 0
        if !this.silent_error
            throw Error(msg, fn, extra)
        this.error_log := 'JSON ERROR`n`nTimestamp:`n' A_Now '`n`nMessage:`n' msg '`n`nFunction:`n' fn '()' (extra = '' ? '' : '`n`nExtra:`n') extra '`n'
        return ''
    }
}