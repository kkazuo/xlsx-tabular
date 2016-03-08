{-# LANGUAGE FlexibleContexts  #-}
{-# LANGUAGE LambdaCase        #-}
{-# LANGUAGE OverloadedStrings #-}
-- | Convinience utility to decode Xlsx tabular.
module Codec.Xlsx.Util.Tabular
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
       , toTableRowsFromFile
       , toTableRows
       , parseJSON
       , toJSON
       ) where

import Debug.Trace

import Codec.Xlsx
import Codec.Xlsx.Formatted
import Codec.Xlsx.Util.Tabular.Types
import Codec.Xlsx.Util.Tabular.Json
import Control.Applicative
import Control.Lens
import Control.Monad (join)
import Data.List (find)
import Data.Map
import Data.Maybe (fromMaybe, isJust)
import Data.Text (Text)

import Data.Aeson (toJSON, parseJSON, encode)

import qualified Data.ByteString.Lazy as ByteString

import Data.IntSet (IntSet)
import qualified Data.IntSet as IntSet


-- |Read from Xlsx file as tabular rows
toTableRowsFromFile :: Int -- ^ Starting row index (header row)
                    -> String -- ^ File name
                    -> IO (Maybe Tabular)
toTableRowsFromFile offset fname = do
  s <- ByteString.readFile fname
  let xlsx = toXlsx s
      rows = toTableRows xlsx (firstSheetName xlsx) offset
  pure rows
  where
    firstSheetName xlsx =
      keys (xlsx ^. xlSheets)
      & head

type Rows =
  [(Int, Cols)]

type Cols =
  [(Int, Cell)]

type RowValues =
  [(Int, [(Int, Maybe CellValue)])]

-- |Decode cells as tabular rows.
toTableRows :: Xlsx -- ^ Xlsx Workbook
            -> Text -- ^ Worksheet name to decode
            -> Int -- ^ Starting row index (header row)
            -> Maybe Tabular
toTableRows xlsx sheetName offset =
  decodeRows <$> styles <*> Just offset <*> rows
  where
    styles = parseStyleSheet (xlsx ^. xlStyles) ^? _Right
    rows =
      xlsx
      ^? ixSheet sheetName
      . wsCells
      . to toRows

decodeRows ss offset rs =
  mkTabular header' rows
  where
    rs' = getCells ss offset rs
    header = head rs' ^. _2
    header' =
      header
      & fmap toText
      & join
    toText (i, Just (CellText t)) = [mkTabularHead i t]
    toText _ = []
    cix = fmap (view tabularHeadIx) header'
      & IntSet.fromList
    rows =
      fmap rowValue (tail rs')
    rowValue rvs =
      mkTabularRow (rvs ^. _1) (rvs ^. _2 & fmap f & join)
      where
        f (i, cell) =
          [cell | cix ^. contains i]

-- |行から値のあるセルを取り出す
getCells :: StyleSheet -- ^スタイルシート
         -> Int -- ^開始行
         -> Rows -- ^セル行
         -> RowValues
getCells ss i rs =
  startAt ss i rs
  & takeContiguous i
  & takeUntil ss
  & fmap rvs
  & filter vs
  where
    rvs (i, cs) =
      (i, rowValues cs)
    filter =
      Prelude.filter
    vs (i, cs) =
      any (\(_, v) -> isJust v) cs

startAt :: StyleSheet -> Int -> Rows -> Rows
startAt ss i rs =
  dropWhile f rs
  where
    f (x, _) =
      x < i

-- |指定の行から連続している行を取り出す
takeContiguous :: Int -> Rows -> Rows
takeContiguous i rs =
  [r | (x, r@(y, _)) <- zip [i..] rs, x == y]

-- |有効セルのすべてに枠線（Bottom側）が存在しなくなる
-- |すなわち枠囲みの欄外になるまでの行を取り出す
takeUntil :: StyleSheet -> Rows -> Rows
takeUntil ss rs =
  takeWhile f rs
  where
    f (i, cs) =
      or $ rowBordersHas borderBottom ss cs

rowBordersHas v ss cs =
  x
  where
    x =
      fmap f cs
    f (i, cell) =
      cellHasBorder ss cell v

rowValues cs =
  x
  where
    x =
      fmap f cs
    f (i, cell) =
      (i, cell ^. cellValue)

cellHasBorder ss cell v =
  fromMaybe False mb
  where
    b = cellBorder ss cell
    mb = borderStyleHasLine v <$> b

cellBorder :: StyleSheet -> Cell -> Maybe Border
cellBorder ss cell =
  view cellStyle cell
  >>= pure . xf
  >>= view cellXfBorderId
  >>= pure . bd
  where
    xf n = (ss ^. styleSheetCellXfs) !! n
    bd n = (ss ^. styleSheetBorders) !! n

borderStyleHasLine v b =
  fromMaybe False value
  where
  value =
    view v b
    >>= view borderStyleLine
    >>= pure . (/= LineStyleNone)
