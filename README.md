# update_price

## Usage

```shell
python3 update_price.py --out ${OUT_PATH} --config ${CONFIG_PATH} ${FILE}
```

The default value of `${CONFIG_PATH}` is update_rules.txt. If not use `--out ${OUT_PATH}`, the modified file will output to `./output`. So, you can use:

```shell
python3 update_price.py ${FILE}
```

For example:

```shell
python3 update_price.py \home\user\Downloads\test.xlsx
```
