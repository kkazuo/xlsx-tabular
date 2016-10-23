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
         -- * Methods
       , def
         -- * Functions
       , toTableRowsFromFile
       , toTableRows
       , toTableRows'
         -- * Custom row predicates
       , toTableRowsCustom
       ) where

import Codec.Xlsx.Util.Tabular.Imports
import qualified Data.ByteString.Lazy as ByteString

type Row =
  (Int, Cols)

type Rows =
  [(Int, Cols)] -- [Row]

type Cols =
  [(Int, Cell)]

type RowValues =
  [(Int, [(Int, Maybe CellValue)])]

-- | A @RowPredicate@ is given the Xlsx "StyleSheet" as well as the
-- row itself (consisting of the row's index and the row's cells) and
-- should return @True@ if the row is part of the table and false
-- otherwise.
type RowPredicate =
  StyleSheet -> Row -> Bool

-- |Read tabular rows from the first sheel of an Xlsx file.
-- The table is assumed to consist of all contiguous rows
-- that have bottom border lines, starting with the header.
toTableRowsFromFile :: Int -- ^ Starting row index (header row)
                    -> String -- ^ File name
                    -> IO (Maybe Tabular)
toTableRowsFromFile offset fname = do
  s <- ByteString.readFile fname
  let xlsx = toXlsx s
      rows = toTableRows' xlsx offset
  pure rows

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
toTableRows' xlsx offset =
  toTableRows xlsx firstSheetName offset
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
  & tabularRows .~ rows
  where
    rs' = getCells p offset rs
    header = head rs' ^. _2
    header' =
      header
      & fmap toText
      & join
    toText (i, Just (CellText t)) = [def
                                     & tabularHeadIx .~ i
                                     & tabularHeadLabel .~ t]
    toText _ = []
    cix = fmap (view tabularHeadIx) header'
      & fromList
    rows =
      fmap rowValue (tail rs')
    rowValue rvs =
      def
      & tabularRowIx .~ (rvs ^. _1)
      & tabularRowCells .~ (rvs ^. _2 & fmap f & join)
      where
        f (i, cell) =
          [cell | cix ^. contains i]

-- |Pickup cells that has value from line
getCells :: (Row -> Bool) -- ^ Predicate
         -> Int -- ^ Start line number
         -> Rows -- ^ cell rows
         -> RowValues
getCells p i rs =
  startAt i rs
  & takeContiguous i
  & takeWhile p
  & fmap rvs
  & filter vs
  where
    rvs (i, cs) =
      (i, rowValues cs)
    filter =
      Prelude.filter
    vs (i, cs) =
      any (\(_, v) -> isJust v) cs

startAt :: Int -> Rows -> Rows
startAt i rs =
  dropWhile f rs
  where
    f (x, _) =
      x < i

-- |Take contiguous rows that start from i
takeContiguous :: Int -> Rows -> Rows
takeContiguous i rs =
  [r | (x, r@(y, _)) <- zip [i..] rs, x == y]

-- |Take rows while all valued cell has bottom border line.
-- |  * no bottom border line means out of table.
borderBottomPredicate :: RowPredicate -- StyleSheet -> Row -> Bool
borderBottomPredicate ss = or . rowBordersHas borderBottom ss . snd

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
