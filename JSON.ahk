; JSON for AutoHotkey
; Copyright (c) 2018 Kurt McKee <contactme@kurtmckee.org>
; The code is licensed under the terms of the MIT license.
; https://github.com/kurtmckee/ahk_json

; VERSION = "1.0"
; -------------------------------------------------------
json_escape(blob)
{
    hexadecimal := "0123456789abcdef"

    escapes := {}
    escapes["`b"] := "\b"
    escapes["`f"] := "\f"
    escapes["`n"] := "\n"
    escapes["`r"] := "\r"
    escapes["`t"] := "\t"
    escapes["/"] := "\/"
    escapes["\"] := "\\"
    escapes[""""] := "\"""


    loop, % strlen(blob)
    {
        character := substr(blob, a_index, 1)
        value := ord(character)

        ; Use simple escapes for reserved characters.
        if (instr("`b`f`n`r`t/\""", character))
        {
            escaped_blob .= escapes[character]
        }

        ; Allow ASCII characters through without modification.
        else if (value >= 32 and value <= 126)
        {
            escaped_blob .= character
        }

        ; Use Unicode escapes for everything else.
        else
        {
            hex1 := substr(hexadecimal, ((value & 0xF000) >> 12) + 1, 1)
            hex2 := substr(hexadecimal, ((value & 0xF00) >> 8) + 1, 1)
            hex3 := substr(hexadecimal, ((value & 0xF0) >> 4) + 1, 1)
            hex4 := substr(hexadecimal, ((value & 0xF) >> 0) + 1, 1)
            escaped_blob .= "\u" . hex1 . hex2 . hex3 . hex4
        }
    }

    return escaped_blob
}

json_unescape(blob)
{
    escapes := {}
    escapes["b"] := "`b"
    escapes["f"] := "`f"
    escapes["n"] := "`n"
    escapes["r"] := "`r"
    escapes["t"] := "`t"
    escapes["/"] := "/"
    escapes["\"] := "\"
    escapes[""""] := """"


    index := 1
    loop
    {
        if (index > strlen(blob))
        {
            break
        }

        character := substr(blob, index, 1)
        next_character := substr(blob, index + 1, 1)
        if (character != "\")
        {
            unescaped_blob .= character
        }
        else if (instr("bfnrt/\""", next_character))
        {
            unescaped_blob .= escapes[next_character]
            index += 1
        }
        else if (next_character == "u")
        {
            unicode_character := chr("0x" . substr(blob, index + 2, 4))
            unescaped_blob .= unicode_character
            index += 5
        }

        index += 1
    }

    return unescaped_blob
}

json_get_object_type(object)
{
    ; Identify the object type and return either "dict" or "list".
    object_type := "list"
    if (object.length() == 0)
    {
        object_type := "dict"
    }
    for key in object
    {
        ; The current AutoHotkey list implementation will loop through its
        ; indexes in order from least to greatest. If the object can be
        ; represented as a list, each key will match the a_index variable.
        ; However, if it is a sparse list (that is, if it has non-consective
        ; list indexes) then it must be represented as a dict.
        if (key != a_index)
        {
            object_type := "dict"
        }
    }

    return object_type
}

toJSON(info)
{
    ; Differentiate between a list and a dictionary.
    object_type := json_get_object_type(info)

    for key, value in info
    {
        ; Only include a key if this is a dictionary.
        if (object_type == "dict")
        {
            escaped_key := json_escape(key)
            blob .= """" . escaped_key . """: "
        }

        if (isobject(value))
        {
            blob .= toJSON(value) . ", "
        }
        else
        {
            escaped_value := json_escape(value)
            blob .= """" . escaped_value . """, "
        }
    }

    ; Remove the final trailing comma.
    if (substr(blob, -1, 2) == ", ")
    {
        blob := substr(blob, 1, -2)
    }

    ; Wrap the string in brackets or braces, as appropriate.
    if (object_type == "list")
    {
        blob := "[" . blob . "]"
    }
    else
    {
        blob := "{" . blob . "}"
    }

    return blob
}

fromJSON(blob)
{
   
   blob_length := strlen(blob)
    index_left := 0
    index_right := 0

    ; Identify the object type.
    loop, % blob_length
    {
        index_left += 1

        if (substr(blob, a_index, 1) == "[")
        {
            object_type := "list"
            info := []
            break
        }
        else if (substr(blob, a_index, 1) == "{")
        {
            object_type := "dict"
            info := {}
            break
        }
    }

    ; Extract all key/value pairs.
    loop, % blob_length
    {
        ; Extract the key.
        ; Use an integer key if this is a list object.
        if (object_type == "list")
        {
            key := info.length() + 1
        }
        else
        {
            ; Find the left side of the key.
            loop, % blob_length
            {
                index_left += 1

                if (substr(blob, index_left, 1) == """")
                {
                    break
                }
            }

            index_right := index_left

            ; Find the right side of the key.
            loop, % blob_length
            {
                index_right += 1

                ; Skip escaped characters, in case they are quotation marks.
                if (substr(blob, index_right, 1) == "\")
                {
                    index_right += 1
                }
                else if (substr(blob, index_right, 1) == """")
                {
                    break
                }
            }

            ; Store the key.
            escaped_key := substr(blob, index_left + 1, index_right - index_left - 1)
            key := json_unescape(escaped_key)
        }

        ; Pass over whitespace and any colons that separate key-value pairs.
        index_left := index_right + 1
        loop, % blob_length
        {
            index_left += 1

            if (not instr("`b`f`n`r`t :", substr(blob, index_left, 1)))
            {
                break
            }
        }

        ; If the value isn't a string, adjust the left index to include
        ; the beginning of the literal, dictionary, or list.
        depth := 0
        in_string := true
        value_type := "str"
        index_right := index_left + 1
        if (substr(blob, index_left, 1) != """")
        {
            in_string := false
            value_type := "literal"
            if (substr(blob, index_left, 1) == "{")
            {
                depth := 1
                value_type := "dict"
            }
            else if (substr(blob, index_left, 1) == "[")
            {
                depth := 1
                value_type := "list"
            }

            index_left -= 1
        }

        ; Find the right edge of the value.
        ;
        ; The loop will isolate the entire value, whether it's a string,
        ; list, dictionary, boolean, integer, float, or null. For example:
        ;
        ;   *   "abc"
        ;   *   123
        ;   *   true
        ;   *   false
        ;   *   null
        ;   *   [123, {"abc": true}]
        ;   *   {"a": [123, null]}
        loop
        {
            if (index_right > blob_length)
            {
                return info
            }

            if (in_string)
            {
                ; If the right index is passing through a string and the
                ; closing quotation mark is encountered, flag that the index
                ; is no longer in a string, and exit the loop if the value is
                ; a string.
                if (substr(blob, index_right, 1) == """")
                {
                    in_string := false
                    if (value_type == "str")
                    {
                        break
                    }
                }
                ; If the right index encounters a backslash in a string, the
                ; next character is guaranteed to still be in the string. Move
                ; the right index forward an additional character in case
                ; the escaped character is a quotation mark.
                else if (substr(blob, index_right, 1) == "\")
                {
                    index_right += 1
                }
            }

            ; If the right index encounters a quotation mark but is not already
            ; in a string, flag that the index is now passing through a string.
            else if (substr(blob, index_right, 1) == """")
            {
                in_string := true
            }

            ; If the value is a dictionary, keep track of the depth of any
            ; nested dictionaries. If the final closing curly brace is found,
            ; move the right index forward so that the right curly brace will
            ; be included in the value and exit the loop.
            else if (value_type == "dict")
            {
                ; If the value is a dictionary
                if (substr(blob, index_right, 1) == "{")
                {
                    depth += 1
                }
                else if (substr(blob, index_right, 1) == "}")
                {
                    depth -= 1
                    index_right += 1
                    if (depth == 0)
                    {
                        break
                    }
                }
            }

            ; If the value is a list, keep track of the depth of any nested
            ; lists. If the final closing bracket is found, move the right
            ; index forward so that the right bracket will be included in
            ; the value and exit the loop.
            else if (value_type == "list")
            {
                if (substr(blob, index_right, 1) == "[")
                {
                    depth += 1
                }
                else if (substr(blob, index_right, 1) == "]")
                {
                    index_right += 1
                    depth -= 1
                    if (depth == 0)
                    {
                        break
                    }
                }
            }

            ; If the value is a literal, such as a boolean or integer, just
            ; watch for any character that will indicate that the end of the
            ; literal has been encountered.
            else if (value_type == "literal")
            {
                if (instr("`b`f`n`r`t ,]}", substr(blob, index_right, 1)))
                {
                    break
                }
            }

            index_right += 1
        }

        ; Extract the value, now that its left and right sides have been found.
        value := substr(blob, index_left + 1, index_right - index_left - 1)

        ; Recursively parse dictionaries and lists.
        if (value_type == "dict" or value_type == "list")
        {
            value := fromJSON(value)
        }
        ; Escape string values.
        else if (value_type == "str")
        {
            value := json_unescape(value)
        }
        ; Convert boolean and null literals to booleans.
        else if (value == "true")
        {
            value := true
        }
        else if (value == "false")
        {
            value := false
        }
        else if (value == "null")
        {
            value := false
        }

        ; Save the key/value pair.
        info[key] := value

        ; Move the index.
        index_left := index_right + 1
    }

    return info
}

isJSONValid(string) 
{
	static doc := ComObjCreate("htmlfile")
		, __ := doc.write("<meta http-equiv=""X-UA-Compatible"" content=""IE=9"">")
		, parse := ObjBindMethod(doc.parentWindow.JSON, "parse")
   try %parse%(string)
   catch
      return false
   return true
}