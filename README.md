# 711-shipping-tools

Tools for shipping items with 7-11 in Taiwan.

Currently focuses on bulk store-to-store shipping, which requires filling in an XLSX template file and uploading it to:

[https://myship2.7-11.com.tw/C2C/Import](https://myship2.7-11.com.tw/C2C/Import)

## `create_bulk_xlsx.py`

Fills in the template XLSX's fields.  Accepts data (1) via TSV from file or stdin, and (2) as command-line arguments for specifying "default fields".  If (1) is not specified, only one row is filled in according to (2).

For example:

```sh
$ ./create_bulk_xlsx.py \
    --verbose \
    --template bulk_template.xlsx \
    --output output.xlsx \
    --sender-name Joel \
    --input - \
    --field receiver_name \
    --field receiver_email << EOF
Some Dude,somedude@hotmail.com
EOF
2024-06-07 15:08:14,871 [DEBUG] Fill row: ['Joel', None, None, None, None, None, 'Some Dude', None, 'somedude@hotmail.com', None, None]
```
