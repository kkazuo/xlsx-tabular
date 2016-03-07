{-# LANGUAGE FlexibleContexts  #-}
{-# LANGUAGE LambdaCase        #-}
{-# LANGUAGE OverloadedStrings #-}

module Codec.Xlsx.Util.Tabular
       ( someFunc
       ) where

import Debug.Trace

import Codec.Xlsx
import Codec.Xlsx.Formatted
import Control.Applicative
import Control.Lens
import Control.Monad (join)
import Data.List (find)
import Data.Map

import qualified Data.ByteString.Lazy as ByteString

import Data.IntSet (IntSet)
import qualified Data.IntSet as IntSet

someFunc :: IO ()
someFunc = do
  s <- ByteString.readFile "sample.xlsx"
  let book = toXlsx s
      sheetNames = keys (book ^. xlSheets)
      sheetName = head sheetNames
      Right styles = parseStyleSheet (book ^. xlStyles)
      Just sheet = book ^? ixSheet sheetName
      rows = toRows (sheet ^. wsCells)
  putStrLn (show (decodeRows styles 9 rows))

type Rows =
  [(Int, Cols)]

type Cols =
  [(Int, Cell)]

type RowValues =
  [(Int, [(Int, Maybe CellValue)])]

decodeRows ss offset rs =
  (header', rows)
  where
    rs' = getCells ss offset rs
    header = (rs' !! 0) ^. _2
    header' =
      header
      & fmap toText
      & join
    toText (i, Just (CellText t)) = [(i, t)]
    toText _ = []
    cix = fmap (view _1) header'
      & IntSet.fromList
    rows =
      fmap rowValue (tail rs')
    rowValue rvs =
      rvs & _2 %~ join . fmap f
      where
        f (i, cell) =
          if cix ^. contains i
          then [cell]
          else []

-- |行から値のあるセルを取り出す
getCells :: StyleSheet -- ^スタイルシート
         -> Int -- ^開始行
         -> Rows -- ^セル行
         -> RowValues
getCells ss i rs =
  startAt ss i rs
  & takeUntil ss
  & fmap rvs
  & filter vs
  where
    rvs (i, cs) =
      (i, rowValues cs)
    filter =
      Prelude.filter
    vs (i, cs) =
      any (\(_, v) -> v /= Nothing) cs

startAt :: StyleSheet -> Int -> Rows -> Rows
startAt ss i rs =
  dropWhile f rs
  where
    f (x, _) =
      x < i

takeUntil :: StyleSheet -> Rows -> Rows
takeUntil ss rs =
  takeWhile f rs
  where
    f (i, cs) =
      and $ rowBordersHas borderBottom ss cs

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
  maybe False id mb
  where
    b = cellBorder ss cell
    mb = borderStyleHasLine v <$> b

cellBorder :: StyleSheet -> Cell -> Maybe Border
cellBorder ss cell =
  view cellStyle cell
  >>= pure . xf
  >>= (view cellXfBorderId)
  >>= pure . bd
  where
    xf n = (ss ^. styleSheetCellXfs) !! n
    bd n = (ss ^. styleSheetBorders) !! n

borderStyleHasLine v b =
  maybe False id value
  where
  value =
    view v b
    >>= (view borderStyleLine)
    >>= pure . (/= LineStyleNone)
