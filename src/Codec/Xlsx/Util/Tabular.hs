{-# LANGUAGE FlexibleContexts  #-}
{-# LANGUAGE LambdaCase        #-}
{-# LANGUAGE OverloadedStrings #-}

{- | Convenience utility to read Xlsx tabular cells.

The majority of the @toTableRows*@ functions assume that the table
of interest consiste of contiguous rows styled with borders lines
surrounding all cells, with possible text above and below the table
that is not of interest. Like so:

@
Some documentation here....
---------------------------
| Header1 | Header2 | ... |
---------------------------
| Value1  | Value2  | ... |
---------------------------
| Value1  | Value2  | ... |
---------------------------
Maybe some annoying text here, I don't care about.
@

The heauristic used for table row selection in these functions is
that any table rows will have a bottom border line.

If the above heuristic is not valid for your table you can instead
provide your own row selection predicate to the `toTableRowsCustom`
function. For example, the predicate @\\_ _ -> True@ (or @(const
. const) True@) will select all contiguous rows.

-}
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
         -- * Functions
       , toTableRowsFromFile
       , toTableRows
       , toTableRows'
         -- * Custom row predicates
       , toTableRowsCustom
       ) where

import Codec.Xlsx.Util.Tabular.Imports
import qualified Data.ByteString.Lazy as ByteString

type Row = (Int, Cols)

type Rows = [(Int, Cols)] -- [Row]

type Cols = [(Int, Cell)]

type RowValues = [(Int, [(Int, Maybe CellValue)])]

-- | A @RowPredicate@ is given the Xlsx "StyleSheet" as well as the
-- row itself (consisting of the row's index and the row's cells) and
-- should return @True@ if the row is part of the table and false
-- otherwise.
type RowPredicate = StyleSheet -> Row -> Bool

-- |Read tabular rows from the first sheel of an Xlsx file.
-- The table is assumed to consist of all contiguous rows
-- that have bottom border lines, starting with the header.
toTableRowsFromFile :: Int -- ^ Starting row index (header row)
                    -> String -- ^ File name
                    -> IO (Maybe Tabular)
toTableRowsFromFile offset fname =
  flip toTableRows' offset . toXlsx <$> ByteString.readFile fname

-- |Decode cells as tabular rows.
-- The table is assumed to consist of all contiguous rows
-- that have bottom border lines, starting with the header.
toTableRows :: Xlsx -- ^ Xlsx Workbook
            -> Text -- ^ Worksheet name to decode
            -> Int -- ^ Starting row index (header row)
            -> Maybe Tabular
toTableRows = toTableRowsCustom borderBottomPredicate

-- |Decode cells from first sheet as tabular rows.
-- The table is assumed to consist of all contiguous rows
-- that have bottom border lines, starting with the header.
toTableRows' :: Xlsx -- ^ Xlsx Workbook
             -> Int -- ^ Starting row index (header row)
             -> Maybe Tabular
toTableRows' xlsx = toTableRows xlsx firstSheetName
  where
    firstSheetName = fst $ head $ xlsx ^. xlSheets
    -- ^ TODO: Is this still true with xlsx-0.3 or are sheets now
    -- in alphabetical order??

-- | Decode cells as tabular rows.
--   The table is assumed to consist of all contiguous rows
--   that fulfill the given predicate, starting with the header.
--
--   The predicate function is given the Xlsx @StyleSheet@ as well
--   as a row (consisting of the row's index and the row's cells)
--   and should return @True@ if the row is part of the table.
--
--   Since 0.1.1
toTableRowsCustom :: (StyleSheet -> (Int, [(Int, Cell)]) -> Bool)
                           -- ^ Predicate for row selection
                  -> Xlsx  -- ^ Xlsx Workbook
                  -> Text  -- ^ Worksheet name to decode
                  -> Int   -- ^ Starting row index (header row)
                  -> Maybe Tabular
toTableRowsCustom predicate xlsx sheetName offset = do
  styles <- parseStyleSheet (xlsx ^. xlStyles) ^? _Right
  rows   <- xlsx ^? ixSheet sheetName . wsCells . to toRows
  decodeRows (predicate styles) offset rows

decodeRows p offset rs = if null rs' then Nothing else Just $
  def
  & tabularHeads .~ header'
  & tabularRows  .~ rows
  where
    rs' = getCells p offset rs
    header = head rs' ^. _2
    header' = join $ map toText header
    toText (i, Just (CellText t)) = [def
                                     & tabularHeadIx .~ i
                                     & tabularHeadLabel .~ t]
    toText _ = []
    cix = fromList $ map (view tabularHeadIx) header'
    rows = fmap rowValue (tail rs')
    rowValue rvs = def
                 & tabularRowIx    .~ fst rvs
                 & tabularRowCells .~ (map snd . filter f . snd) rvs
      where
        f = flip member cix . fst

-- |Pickup cells that has value from line
getCells :: (Row -> Bool) -- ^ Predicate
         -> Int -- ^ Start line number
         -> Rows -- ^ cell rows
         -> RowValues
getCells p i = filter (any (isJust . snd) . snd)
              . (fmap . fmap) rowValues
              . takeContiguous i
              . takeWhile p
              . startAt i

startAt :: Int -> Rows -> Rows
startAt i = dropWhile ((< i) . fst)

-- |Take contiguous rows that start from i
takeContiguous :: Int -> Rows -> Rows
--takeContiguous i rs = [r | (x, r@(y, _)) <- zip [i..] rs, x == y]
takeContiguous i = map snd . filter (uncurry (==) . fmap fst) . zip [i..]

rowValues = map (fmap _cellValue)

-- Predicate for at least one cell having a bottom border style.

-- |Take rows while all valued cell has bottom border line.
-- |  * no bottom border line means out of table.
borderBottomPredicate :: RowPredicate -- StyleSheet -> Row -> Bool
borderBottomPredicate ss = or . rowBordersHas borderBottom ss . snd

rowBordersHas v ss = map (cellHasBorder v ss . snd)


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
