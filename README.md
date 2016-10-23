# xlsx-tabular
Xlsx table decode utility

[![packagename on Stackage LTS 3](http://stackage.org/package/xlsx-tabular/badge/lts)](http://stackage.org/lts/package/xlsx-tabular)

## Library Usage Example

```haskell
#!/usr/bin/env stack
-- stack --resolver=lts-5.9 runghc --package=xlsx-tabular

import Data.Aeson
import qualified Data.ByteString.Lazy.Char8 as BS8
import Codec.Xlsx.Util.Tabular
import Codec.Xlsx.Util.Tabular.Json

main = do
  r <- toTableRowsFromFile 8 "your-sample.xlsx"
  let json = encode r
  BS8.putStrLn json
```


## Contributors

  * [Björn Buckwalter](https://github.com/bjornbm), who has extended this library's core usability. (```toTableRowsCustom```)
  
