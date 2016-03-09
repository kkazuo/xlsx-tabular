{-# LANGUAGE TemplateHaskell   #-}

module Codec.Xlsx.Util.Tabular.Types
       ( Tabular
       , mkTabular
       , tabularHeads
       , tabularRows
       , TabularHead
       , mkTabularHead
       , tabularHeadIx
       , tabularHeadLabel
       , TabularRow
       , mkTabularRow
       , tabularRowIx
       , tabularCells
       ) where

import Codec.Xlsx (CellValue)
import Control.Lens (makeLenses)
import Data.Text (Text)


-- | Tabular cells
data Tabular = Tabular
  { _tabularHeads :: [TabularHead]
  , _tabularRows :: [TabularRow]
  }
  deriving (Show)

-- | Tabular header
data TabularHead = TabularHead
  { _tabularHeadIx :: Int -- ^ Column index
  , _tabularHeadLabel :: Text -- ^ Column label
  }
  deriving (Show)

-- | Tabular row
data TabularRow = TabularRow
  { _tabularRowIx :: Int -- ^ Row index
  , _tabularCells :: [Maybe CellValue] -- ^ Row values
  }
  deriving (Show)


makeLenses ''Tabular
makeLenses ''TabularHead
makeLenses ''TabularRow


mkTabular :: [TabularHead]
          -> [TabularRow]
          -> Tabular
mkTabular =
  Tabular


mkTabularHead :: Int
              -> Text
              -> TabularHead
mkTabularHead =
  TabularHead


mkTabularRow :: Int
             -> [Maybe CellValue]
             -> TabularRow
mkTabularRow =
  TabularRow
