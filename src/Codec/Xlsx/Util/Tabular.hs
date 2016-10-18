{-# LANGUAGE FlexibleContexts  #-}
{-# LANGUAGE LambdaCase        #-}
{-# LANGUAGE OverloadedStrings #-}
-- | Convinience utility to read Xlsx tabular cells.
module Codec.Xlsx.Util.Tabular
       (
         -- * Types
         Tabular
       , TabularHead
       , TabularRow
         -- * Lenses
         -- ** Tabular
       , tabularHeads
       , tabularRows
         -- ** TabularHead
       , tabularHeadIx
       , tabularHeadLabel
         -- ** TabularRow
       , tabularRowIx
       , tabularRowCells
         -- * Methods
       , def
         -- * Functions
       , toTableRowsFromFile
       , toTableRows
       , toTableRows'
       ) where

import Codec.Xlsx.Util.Tabular.Imports
import qualified Data.ByteString.Lazy as ByteString

type Rows = [(Int, Cols)]

type Cols = [(Int, Cell)]

type RowValues = [(Int, [(Int, Maybe CellValue)])]


-- |Read from Xlsx file as tabular rows
toTableRowsFromFile :: Int -- ^ Starting row index (header row)
                    -> String -- ^ File name
                    -> IO (Maybe Tabular)
toTableRowsFromFile offset fname =
  flip toTableRows' offset . toXlsx <$> ByteString.readFile fname

-- |Decode cells as tabular rows.
toTableRows :: Xlsx -- ^ Xlsx Workbook
            -> Text -- ^ Worksheet name to decode
            -> Int -- ^ Starting row index (header row)
            -> Maybe Tabular
toTableRows xlsx sheetName offset =
  decodeRows <$> styles <*> Just offset <*> rows
  where
    styles = parseStyleSheet (xlsx ^. xlStyles) ^? _Right
    rows = xlsx ^? ixSheet sheetName . wsCells . to toRows

-- |Decode cells as tabular rows from first sheet.
toTableRows' :: Xlsx -- ^ Xlsx Workbook
             -> Int -- ^ Starting row index (header row)
             -> Maybe Tabular
toTableRows' xlsx = toTableRows xlsx firstSheetName
  where
    firstSheetName = fst $ head $ xlsx ^. xlSheets

decodeRows ss offset rs =
  def
  & tabularHeads .~ header'
  & tabularRows .~ rows
  where
    rs' = getCells ss offset rs
    header = head rs' ^. _2
    header' = join $ toText <$> header
    toText (i, Just (CellText t)) = [def
                                     & tabularHeadIx .~ i
                                     & tabularHeadLabel .~ t]
    toText _ = []
    cix = fmap (view tabularHeadIx) header' & fromList
    rows = fmap rowValue (tail rs')
    rowValue rvs = def
                 & tabularRowIx .~ (rvs ^. _1)
                 & tabularRowCells .~ (rvs ^. _2 & fmap f & join)
      where
        f (i, cell) = [cell | cix ^. contains i]

-- |行から値のあるセルを取り出す
getCells :: StyleSheet -- ^スタイルシート
         -> Int -- ^開始行
         -> Rows -- ^セル行
         -> RowValues
getCells ss i = filter (any (isJust . snd) . snd)
              . (fmap . fmap) rowValues
              . takeContiguous i
              . takeUntil ss
              . startAt ss i

startAt :: StyleSheet -> Int -> Rows -> Rows
startAt ss i = dropWhile ((< i) . fst)

-- |指定の行から連続している行を取り出す
takeContiguous :: Int -> Rows -> Rows
takeContiguous i rs = [r | (x, r@(y, _)) <- zip [i..] rs, x == y]

-- |有効セルのすべてに枠線（Bottom側）が存在しなくなる
-- |すなわち枠囲みの欄外になるまでの行を取り出す
takeUntil :: StyleSheet -> Rows -> Rows
takeUntil ss = takeWhile f
  where
    f (i, cs) = or $ rowBordersHas borderBottom ss cs

rowBordersHas v ss = fmap (cellHasBorder v ss . snd)

rowValues = fmap (fmap (view cellValue))

cellHasBorder v ss cell = fromMaybe False mb
  where
    mb = borderStyleHasLine v <$> cellBorder ss cell

cellBorder :: StyleSheet -> Cell -> Maybe Border
cellBorder ss cell = fmap xf (view cellStyle cell)
                 >>= fmap bd . view cellXfBorderId
  where
    xf n = (ss ^. styleSheetCellXfs) !! n
    bd n = (ss ^. styleSheetBorders) !! n

borderStyleHasLine v b = fromMaybe False value
  where
    value = view v b >>= fmap (/= LineStyleNone) . view borderStyleLine
