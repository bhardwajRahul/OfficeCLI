// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests for the shared Units conversion class, verifying precision of
/// twips→pt and EMU→pt conversions used in HTML preview rendering.
/// </summary>
public class UnitsTests
{
    // ==================== TwipsToPt ====================

    [Fact]
    public void TwipsToPt_ExactConversion_1pt()
    {
        Units.TwipsToPt(20).Should().Be(1.0);
    }

    [Fact]
    public void TwipsToPt_ExactConversion_72pt()
    {
        // 1 inch = 1440 twips = 72pt
        Units.TwipsToPt(1440).Should().Be(72.0);
    }

    [Fact]
    public void TwipsToPt_HalfPoint()
    {
        Units.TwipsToPt(10).Should().Be(0.5);
    }

    [Fact]
    public void TwipsToPt_Zero()
    {
        Units.TwipsToPt(0).Should().Be(0.0);
    }

    [Fact]
    public void TwipsToPt_String_ValidInput()
    {
        Units.TwipsToPt("240").Should().Be(12.0);
    }

    [Fact]
    public void TwipsToPt_String_InvalidInput()
    {
        Units.TwipsToPt("abc").Should().Be(0.0);
    }

    [Fact]
    public void TwipsToPtStr_FormatsCorrectly()
    {
        Units.TwipsToPtStr("240").Should().Be("12pt");
    }

    [Fact]
    public void TwipsToPtStr_FormatsDecimal()
    {
        // 700 twips = 35pt
        Units.TwipsToPtStr("700").Should().Be("35pt");
    }

    // ==================== EmuToPt ====================

    [Fact]
    public void EmuToPt_ExactConversion_1pt()
    {
        // 1 pt = 12700 EMU
        Units.EmuToPt(12700).Should().Be(1.0);
    }

    [Fact]
    public void EmuToPt_ExactConversion_72pt()
    {
        // 1 inch = 914400 EMU = 72pt
        Units.EmuToPt(914400).Should().Be(72.0);
    }

    [Fact]
    public void EmuToPt_Zero()
    {
        Units.EmuToPt(0).Should().Be(0.0);
    }

    [Fact]
    public void EmuToPt_StandardSlideWidth()
    {
        // Standard widescreen 13.333" slide: 12192000 EMU = 960 pt
        Units.EmuToPt(12192000).Should().Be(960.0);
    }

    [Fact]
    public void EmuToPt_StandardSlideHeight()
    {
        // Standard 7.5" slide height: 6858000 EMU = 540 pt
        Units.EmuToPt(6858000).Should().Be(540.0);
    }

    [Fact]
    public void EmuToPt_SmallValue_NoLoss()
    {
        // 100000 EMU — previously EmuToCm gave 0.278 (truncated from 0.27777...)
        // In pt: 100000 / 12700 = 7.874015... → rounds to 7.87
        Units.EmuToPt(100000).Should().Be(7.87);
    }

    // ==================== HalfPointsToPt ====================

    [Fact]
    public void HalfPointsToPt_ExactConversion()
    {
        Units.HalfPointsToPt(24).Should().Be(12.0);
    }

    [Fact]
    public void HalfPointsToPt_OddValue()
    {
        Units.HalfPointsToPt(21).Should().Be(10.5);
    }

    // ==================== Precision Comparisons ====================

    [Fact]
    public void EmuToPt_MorePreciseThanEmuToCm()
    {
        // Demonstrate that pt conversion preserves more precision than cm.
        // 457200 EMU = 36pt exactly (0.5 inch).
        // In cm: 457200 / 360000 = 1.27 (exact)
        // In pt: 457200 / 12700 = 36.0 (exact)
        // Both are exact here, but for non-round values:

        // 355600 EMU: cm = 0.9877... (rounded to 0.988), pt = 28.0 (exact!)
        Units.EmuToPt(355600).Should().Be(28.0);
    }

    [Fact]
    public void TwipsToPt_MorePreciseThanTwipsToPx()
    {
        // 700 twips: old px = Round(700/1440*96, 1) = 46.7px (truncated)
        // new pt = 700/20 = 35.0pt (exact!)
        Units.TwipsToPt(700).Should().Be(35.0);
    }
}
