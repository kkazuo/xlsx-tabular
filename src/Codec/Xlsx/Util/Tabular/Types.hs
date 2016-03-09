{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE TemplateHaskell   #-}

module Codec.Xlsx.Util.Tabular.Types
       ( Tabular
       , tabularHeads
       , tabularRows
       , TabularHead
       , tabularHeadIx
       , tabularHeadLabel
       , TabularRow
       , tabularRowIx
       , tabularCells
       , def
       ) where

import Codec.Xlsx (CellValue)
import Control.Lens (makeLenses)
import Data.Default (Default, def)
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


instance Default Tabular where
  def = Tabular [] []


instance Default TabularRow where
  def = TabularRow 0 []


instance Default TabularHead where
  def = TabularHead 0 ""
